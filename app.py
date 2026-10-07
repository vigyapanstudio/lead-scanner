"""Vigyapan Lead Scanner - a 3-step website: upload Excel -> scan -> download report."""
import datetime as dt
import hmac
import os

import pandas as pd
import streamlit as st

import auditor
from report import build_report

st.set_page_config(page_title="Vigyapan Lead Scanner", page_icon="🔎", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
[data-testid="stSidebar"], [data-testid="collapsedControl"]{display:none}
.block-container{padding-top:2.5rem;max-width:860px}
h1{color:#1F2A44;margin-bottom:0}
.sub{color:#6B7280;margin:.3rem 0 1.8rem;font-size:1.05rem}
.step{font-weight:700;color:#E8772E;letter-spacing:.06em;font-size:.85rem;text-transform:uppercase;margin-top:1.6rem}
.stButton button, .stDownloadButton button{font-size:1.1rem;padding:.7rem 1.6rem;border-radius:10px}
div[data-testid="stMetricValue"]{color:#1F2A44}
button[kind="primary"]{background:#E8772E !important;border-color:#E8772E !important;color:#fff !important}
button[kind="primary"]:hover{background:#C9601D !important;border-color:#C9601D !important}
.foot{color:#9CA3AF;font-size:.8rem;margin-top:3rem;text-align:center}
</style>
""", unsafe_allow_html=True)


def secret(name: str, default: str = "") -> str:
    """Settings are kept on the server (Streamlit 'Secrets'), never shown to the client."""
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except Exception:
        return os.environ.get(name, default)


st.markdown("# 🔎 Vigyapan Lead Scanner")
st.markdown('<p class="sub">Upload your contact list. We check every company\'s website, Google '
            'visibility (SEO) and Instagram, and give you a ready list of who needs what.</p>',
            unsafe_allow_html=True)

# ---------------- password (only if one is set in Secrets)
PASSWORD = secret("APP_PASSWORD")
if PASSWORD and not st.session_state.get("ok"):
    pw = st.text_input("Enter your password", type="password")
    if st.button("Open", type="primary"):
        if hmac.compare_digest(pw.strip(), PASSWORD):
            st.session_state["ok"] = True
            st.rerun()
        else:
            st.error("Wrong password. Please check and try again.")
    st.stop()

# ---------------- step 1: upload
st.markdown('<div class="step">Step 1 · Upload your Excel file</div>', unsafe_allow_html=True)
up = st.file_uploader("Drag your Excel file here, or click Browse", type=["xlsx", "xls", "csv"])
if not up:
    st.info("Your file just needs columns like **Company, Name, Phone, Email, Website**. "
            "An Instagram column is optional.")
    st.markdown('<div class="foot">Made by Vigyapan Studio</div>', unsafe_allow_html=True)
    st.stop()

if st.session_state.get("file_id") != up.file_id:      # new file -> clear old results
    st.session_state.pop("results", None)
    st.session_state["file_id"] = up.file_id

sheet = 0
try:
    if not up.name.lower().endswith(".csv"):
        xl = pd.ExcelFile(up)
        if len(xl.sheet_names) > 1:
            sheet = st.selectbox("Your file has more than one sheet. Which one has the contacts?",
                                 xl.sheet_names)
        up.seek(0)
    df = auditor.read_database(up, sheet)
except Exception:
    st.error("Sorry, we couldn't open this file. Please save it as a normal Excel file (.xlsx) and try again.")
    st.stop()

MAX_ROWS = int(secret("MAX_ROWS", "3000"))
if len(df) > MAX_ROWS:
    st.warning(f"This file has {len(df)} contacts. Only the first {MAX_ROWS} will be scanned - "
               "please split bigger lists into parts.")
    df = df.head(MAX_ROWS)

guess = auditor.detect_columns(df.columns)
found = [auditor.FIELD_LABELS[f] for f, c in guess.items() if c]
st.success(f"✅ Found **{len(df)} contacts**. Columns we understood: {', '.join(found) or 'none'}.")

# Column fixing is hidden unless something important is missing
needs_fix = not guess["website"] and not guess["email"]
with st.expander("Columns look wrong? Fix them here", expanded=needs_fix):
    options = ["— not in my file —"] + list(df.columns)
    mapping = {}
    cols = st.columns(2)
    for i, f in enumerate(auditor.FIELDS):
        g = guess.get(f)
        sel = cols[i % 2].selectbox(auditor.FIELD_LABELS[f], options,
                                    index=options.index(g) if g in options else 0, key=f"map_{f}")
        mapping[f] = None if sel == options[0] else sel
if not mapping["website"] and not mapping["email"]:
    st.error("We couldn't find a **Website** or **Email** column. Please pick them above.")
    st.stop()

# ---------------- step 2: scan
st.markdown('<div class="step">Step 2 · Start the scan</div>', unsafe_allow_html=True)
workers = int(secret("WORKERS", "10"))
mins = max(1, round(len(df) * 4 / workers / 60))
st.caption(f"This will take about {mins} minute(s). Please keep this page open until it finishes.")

if st.button("🔍 Start scan", type="primary", width="stretch"):
    bar = st.progress(0.0, text="Starting…")
    note = st.empty()

    def progress(done, total, label):
        bar.progress(done / total, text=f"Checked {done} of {total} · {label}")

    opts = {"apify_token": secret("APIFY_TOKEN"), "pagespeed_key": secret("PAGESPEED_KEY"),
            "domain_age": secret("DOMAIN_AGE", "yes").lower() != "no", "workers": workers}
    res = auditor.scan_all(df, mapping, opts, progress, lambda m: note.info(m))
    st.session_state["results"] = res
    st.session_state["report"] = build_report(res, up.name)
    st.session_state["src"] = up.name
    bar.progress(1.0, text="Done!")
    note.empty()
    st.balloons()

res = st.session_state.get("results")
if res is None:
    st.stop()

# ---------------- step 3: results
st.markdown('<div class="step">Step 3 · Download your report</div>', unsafe_allow_html=True)
base = os.path.splitext(st.session_state.get("src", "list"))[0]
st.download_button("⬇️ Download Excel report", st.session_state["report"], type="primary",
                   width="stretch",
                   file_name=f"{base} - Lead Report {dt.datetime.now():%d-%b-%Y}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

opp = res["Opportunities"].fillna("")
a, b, c = st.columns(3)
a.metric("Contacts checked", len(res))
b.metric("🔥 HOT leads", int((res["Lead Priority"] == "HOT").sum()))
c.metric("No website", int(opp.str.contains("New Website").sum()))
d, e, f = st.columns(3)
d.metric("Website needs redesign", int(opp.str.contains("Redesign|Upgrade").sum()))
e.metric("Needs SEO", int(opp.str.contains("SEO").sum()))
f.metric("Instagram leads", int(opp.str.contains("Instagram").sum()))

st.markdown("**Quick look** (full details are in the Excel report)")
show = [c for c in ["Lead Priority", "Company", "Contact Person", "Phone", "Website Status",
                    "Opportunities", "Pitch Note"] if c in res.columns]
order = res["Lead Priority"].map({"HOT": 0, "WARM": 1, "LOW": 2})
st.dataframe(res.assign(_o=order).sort_values("_o")[show], hide_index=True, width="stretch", height=420)
st.markdown('<div class="foot">Made by Vigyapan Studio</div>', unsafe_allow_html=True)
