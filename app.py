"""Vigyapan Lead Scanner - a 3-step website: upload Excel -> scan -> download report."""
import datetime as dt
import hmac
import os

import pandas as pd
import streamlit as st

import auditor
from pdf_report import build_pdf
from report import build_report

st.set_page_config(page_title="Vigyapan Lead Scanner", page_icon="🔎", layout="centered",
                   initial_sidebar_state="collapsed")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@500;600;700;800&family=DM+Sans:wght@400;500;700&display=swap');
html, body, [class*="css"], .stMarkdown, p, label, input, button {font-family:'DM Sans',sans-serif}
[data-testid="stSidebar"], [data-testid="collapsedControl"], header[data-testid="stHeader"]{display:none}
.stApp{background:radial-gradient(circle at 0% 0%,#FFE9DA 0,transparent 35%),
       radial-gradient(circle at 100% 20%,#E6E2FF 0,transparent 40%),
       radial-gradient(circle at 50% 100%,#DDF6F0 0,transparent 45%),#FBFAFF}
.block-container{padding-top:1.6rem;max-width:900px}
/* hero */
.hero{position:relative;overflow:hidden;border-radius:26px;padding:34px 34px 30px;color:#fff;
      background:linear-gradient(120deg,#3B2FC9 0%,#7C3AED 45%,#E8772E 100%);
      box-shadow:0 18px 40px -18px rgba(76,45,200,.55);margin-bottom:22px}
.hero:before,.hero:after{content:"";position:absolute;border-radius:50%;background:rgba(255,255,255,.12)}
.hero:before{width:260px;height:260px;right:-70px;top:-90px}
.hero:after{width:140px;height:140px;right:120px;bottom:-70px;background:rgba(255,200,120,.25)}
.hero .tag{display:inline-block;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.35);
      padding:4px 12px;border-radius:999px;font-size:.78rem;letter-spacing:.08em;font-weight:700}
.hero h1{font-family:'Poppins',sans-serif;font-weight:800;font-size:2.5rem;line-height:1.1;color:#fff;margin:.6rem 0 .4rem;padding:0}
.hero p{font-size:1.06rem;opacity:.93;max-width:560px;margin:0;color:#fff}
/* feature cards */
.feats{display:grid;grid-template-columns:repeat(4,1fr);gap:12px;margin:4px 0 26px}
.feat{border-radius:18px;padding:16px 14px;background:#fff;box-shadow:0 6px 18px -10px rgba(30,30,80,.25);
      border-top:5px solid var(--c)}
.feat .ic{font-size:1.5rem}
.feat b{display:block;font-family:'Poppins',sans-serif;color:#1F2A44;margin:.35rem 0 .15rem;font-size:.98rem}
.feat span{color:#5B6475;font-size:.84rem;line-height:1.35}
@media(max-width:700px){.feats{grid-template-columns:repeat(2,1fr)}.hero h1{font-size:1.9rem}}
/* steps */
.step{display:flex;align-items:center;gap:12px;margin:26px 0 10px;font-family:'Poppins',sans-serif;
      font-weight:700;font-size:1.15rem;color:#1F2A44}
.step .n{width:36px;height:36px;border-radius:12px;display:grid;place-items:center;color:#fff;font-size:1rem;
      background:var(--c);box-shadow:0 6px 14px -6px var(--c)}
/* uploader */
[data-testid="stFileUploaderDropzone"]{background:#fff;border:2.5px dashed #A78BFA;border-radius:18px;padding:26px}
[data-testid="stFileUploaderDropzone"]:hover{border-color:#E8772E;background:#FFF8F2}
/* buttons */
.stButton button,.stDownloadButton button{font-size:1.08rem;padding:.75rem 1.6rem;border-radius:14px;font-weight:700}
button[kind="primary"]{background:linear-gradient(100deg,#E8772E,#F43F5E) !important;border:0 !important;color:#fff !important;
      box-shadow:0 10px 22px -10px rgba(244,63,94,.7)}
button[kind="primary"]:hover{filter:brightness(1.07);transform:translateY(-1px)}
button[kind="secondary"]{background:#fff !important;border:2px solid #7C3AED !important;color:#5B21B6 !important}
[data-testid="stProgress"] div[role="progressbar"] > div > div{background:linear-gradient(90deg,#7C3AED,#E8772E) !important}
/* result tiles */
.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:18px 0 8px}
.tile{border-radius:18px;padding:16px 18px;color:#fff;background:var(--g);box-shadow:0 10px 22px -14px rgba(0,0,0,.45)}
.tile .v{font-family:'Poppins',sans-serif;font-size:2.1rem;font-weight:800;line-height:1}
.tile .l{font-size:.88rem;opacity:.95;margin-top:6px;font-weight:500}
@media(max-width:700px){.tiles{grid-template-columns:repeat(2,1fr)}}
.found{background:#fff;border-radius:16px;padding:14px 18px;border-left:6px solid #10B981;
      box-shadow:0 6px 18px -12px rgba(0,0,0,.3);margin:12px 0;color:#1F2A44}
.hint{background:#fff;border-radius:16px;padding:14px 18px;color:#4B5563;box-shadow:0 6px 18px -12px rgba(0,0,0,.25)}
.pill{display:inline-block;background:#EEF2FF;color:#4338CA;border-radius:999px;padding:2px 10px;margin:2px;font-size:.82rem;font-weight:700}
[data-testid="stExpander"]{background:#fff;border-radius:14px}
.foot{color:#8B90A0;font-size:.82rem;margin-top:3rem;text-align:center}
.foot b{background:linear-gradient(90deg,#7C3AED,#E8772E);-webkit-background-clip:text;color:transparent}
</style>
""", unsafe_allow_html=True)


def step(n, text, color):
    st.markdown(f'<div class="step"><span class="n" style="--c:{color}">{n}</span>{text}</div>',
                unsafe_allow_html=True)


FOOT = '<div class="foot">Made with care by <b>Vigyapan Studio</b></div>'


def secret(name: str, default: str = "") -> str:
    """Settings are kept on the server (Streamlit 'Secrets'), never shown to the client."""
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except Exception:
        return os.environ.get(name, default)


st.markdown("""
<div class="hero">
  <span class="tag">VIGYAPAN STUDIO</span>
  <h1>Lead Scanner</h1>
  <p>Upload your contact list. We check every company's website, Google ranking and Instagram,
  then hand you a ready list of who needs what.</p>
</div>
<div class="feats">
  <div class="feat" style="--c:#3B82F6"><div class="ic">🌐</div><b>Website</b><span>Missing, broken, old or not mobile-friendly</span></div>
  <div class="feat" style="--c:#10B981"><div class="ic">📈</div><b>Google / SEO</b><span>Can customers find them on Google?</span></div>
  <div class="feat" style="--c:#EC4899"><div class="ic">📸</div><b>Instagram</b><span>Do they have a page? Is it active?</span></div>
  <div class="feat" style="--c:#F59E0B"><div class="ic">💬</div><b>Pitch ready</b><span>What to sell each lead, in one line</span></div>
</div>
""", unsafe_allow_html=True)

# ---------------- password (only if one is set in Secrets)
PASSWORD = secret("APP_PASSWORD")
if PASSWORD and not st.session_state.get("ok"):
    st.markdown('<div class="hint">🔒 This scanner is private. Please enter the password you were given.</div>',
                unsafe_allow_html=True)
    pw = st.text_input("Password", type="password")
    if st.button("Open the scanner", type="primary"):
        if hmac.compare_digest(pw.strip(), PASSWORD):
            st.session_state["ok"] = True
            st.rerun()
        else:
            st.error("Wrong password. Please check and try again.")
    st.stop()

# ---------------- step 1: upload
step(1, "Upload your Excel file", "#E8772E")
up = st.file_uploader("Drag your Excel file here, or click Browse", type=["xlsx", "xls", "csv"])
if not up:
    st.markdown('<div class="hint">💡 Your file just needs columns like <span class="pill">Company</span>'
                '<span class="pill">Name</span><span class="pill">Phone</span><span class="pill">Email</span>'
                '<span class="pill">Website</span>. An Instagram column is optional.</div>', unsafe_allow_html=True)
    st.markdown(FOOT, unsafe_allow_html=True)
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
pills = "".join(f'<span class="pill">{f}</span>' for f in found) or "none"
st.markdown(f'<div class="found">✅ Found <b>{len(df)} contacts</b>. Columns we understood: {pills}</div>',
            unsafe_allow_html=True)

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
step(2, "Start the scan", "#7C3AED")
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
    bar.progress(1.0, text="Preparing your PDF report…")
    st.session_state["pdf"] = build_pdf(res, up.name)
    st.session_state["report"] = build_report(res, up.name)
    st.session_state["src"] = up.name
    bar.progress(1.0, text="Done!")
    note.empty()
    st.balloons()

res = st.session_state.get("results")
if res is None:
    st.stop()

# ---------------- step 3: results
step(3, "Download your report", "#10B981")
base = os.path.splitext(st.session_state.get("src", "list"))[0]
stamp = f"{dt.datetime.now():%d-%b-%Y}"
st.download_button("⬇️ Download PDF report", st.session_state["pdf"], type="primary", width="stretch",
                   file_name=f"{base} - Lead Report {stamp}.pdf", mime="application/pdf")
st.download_button("Download as Excel (to sort and filter)", st.session_state["report"], width="stretch",
                   file_name=f"{base} - Lead Report {stamp}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

opp = res["Opportunities"].fillna("")
tiles = [
    ("Contacts checked", len(res), "linear-gradient(135deg,#3B2FC9,#6366F1)"),
    ("🔥 HOT leads", int((res["Lead Priority"] == "HOT").sum()), "linear-gradient(135deg,#E11D48,#F97316)"),
    ("No website", int(opp.str.contains("New Website").sum()), "linear-gradient(135deg,#2563EB,#06B6D4)"),
    ("Website needs redesign", int(opp.str.contains("Redesign|Upgrade").sum()), "linear-gradient(135deg,#7C3AED,#C026D3)"),
    ("Needs SEO", int(opp.str.contains("SEO").sum()), "linear-gradient(135deg,#059669,#10B981)"),
    ("Instagram leads", int(opp.str.contains("Instagram").sum()), "linear-gradient(135deg,#DB2777,#F472B6)"),
]
st.markdown('<div class="tiles">' + "".join(
    f'<div class="tile" style="--g:{g}"><div class="v">{v}</div><div class="l">{l}</div></div>'
    for l, v, g in tiles) + "</div>", unsafe_allow_html=True)

st.markdown('<div class="step" style="font-size:1rem;margin-top:18px">👀 Quick look '
            '<span style="font-weight:500;color:#6B7280;font-size:.88rem">- every contact has its own '
            'detailed card in the PDF</span></div>', unsafe_allow_html=True)
show = [c for c in ["Lead Priority", "Company", "Contact Person", "Phone", "Website Status",
                    "Opportunities", "Pitch Note"] if c in res.columns]
order = res["Lead Priority"].map({"HOT": 0, "WARM": 1, "CHECK": 2, "LOW": 3})
view = res.assign(_o=order).sort_values("_o")[show].reset_index(drop=True)
PRIO_CSS = {"HOT": "background-color:#FFE4E6;color:#BE123C;font-weight:700",
            "WARM": "background-color:#FEF3C7;color:#92400E;font-weight:700",
            "LOW": "background-color:#D1FAE5;color:#065F46;font-weight:700",
            "CHECK": "background-color:#E5E7EB;color:#374151;font-weight:700"}
styled = view.style.map(lambda v: PRIO_CSS.get(v, ""), subset=["Lead Priority"])
st.dataframe(styled, hide_index=True, width="stretch", height=min(460, 36 * (len(view) + 1) + 4))
st.markdown(FOOT, unsafe_allow_html=True)
