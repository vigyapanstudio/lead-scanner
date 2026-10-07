"""Where the lead tracker is saved.

* If a Google Sheet is connected in Secrets  -> saved there permanently (recommended).
* Otherwise                                  -> saved in a file on the server, which is wiped whenever
                                                the website restarts (fine for trying it out)."""
from __future__ import annotations

import datetime as dt
import hashlib
import os
import re

import pandas as pd

LEAD_COLS = ["Lead ID", "Event", "Scanned On", "Company", "Contact Person", "Phone", "Email", "City", "Website",
             "Lead Priority", "Opportunities", "Coverage Prospect", "Pitch Note", "WhatsApp Link",
             "Status", "Follow-up Date", "Notes", "Updated"]
STATUSES = ["New", "Contacted", "Interested", "Follow up", "Not interested", "Won"]
USER_COLS = ["Status", "Follow-up Date", "Notes"]          # what the team edits - never overwritten by a re-scan
LOCAL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "leads.csv")


def lead_id(event: str, company: str, phone: str, email: str) -> str:
    key = "|".join(re.sub(r"\W", "", str(x).lower()) for x in (event, company, phone or email))
    return hashlib.md5(key.encode()).hexdigest()[:10]


class Store:
    def __init__(self, secrets=None):
        self.kind = "local"
        self.error = ""
        self._ws = None
        try:
            if secrets is not None and "gcp_service_account" in secrets and secrets.get("GSHEET_ID"):
                import gspread
                gc = gspread.service_account_from_dict(dict(secrets["gcp_service_account"]))
                sh = gc.open_by_key(str(secrets["GSHEET_ID"]))
                try:
                    self._ws = sh.worksheet("Leads")
                except gspread.WorksheetNotFound:
                    self._ws = sh.add_worksheet("Leads", rows=1000, cols=len(LEAD_COLS))
                self.kind = "google"
        except Exception as e:  # fall back to local file, but tell the user
            self.error = f"Could not open the Google Sheet ({type(e).__name__}). Using temporary storage."
            self.kind = "local"

    # ---------------------------------------------------------------- read / write
    def load(self) -> pd.DataFrame:
        if self.kind == "google":
            rows = self._ws.get_all_values()
            if not rows:
                return pd.DataFrame(columns=LEAD_COLS)
            df = pd.DataFrame(rows[1:], columns=rows[0])
        elif os.path.exists(LOCAL_FILE):
            df = pd.read_csv(LOCAL_FILE, dtype=str).fillna("")
        else:
            df = pd.DataFrame(columns=LEAD_COLS)
        for c in LEAD_COLS:
            if c not in df.columns:
                df[c] = ""
        return df[LEAD_COLS].fillna("").astype(str)

    def save(self, df: pd.DataFrame) -> None:
        df = df.copy()
        for c in LEAD_COLS:
            if c not in df.columns:
                df[c] = ""
        df = df[LEAD_COLS].fillna("").astype(str)
        if self.kind == "google":
            self._ws.clear()
            self._ws.update([LEAD_COLS] + df.values.tolist(), value_input_option="RAW")
        else:
            os.makedirs(os.path.dirname(LOCAL_FILE), exist_ok=True)
            df.to_csv(LOCAL_FILE, index=False)

    # ---------------------------------------------------------------- helpers
    def add_scan(self, results: pd.DataFrame, event: str) -> int:
        """Adds / refreshes leads from a scan. Keeps the team's Status / Follow-up / Notes."""
        now = f"{dt.datetime.now():%Y-%m-%d %H:%M}"
        old = self.load()
        old_by_id = {r["Lead ID"]: r for r in old.to_dict("records")}
        new_rows = []
        for r in results.to_dict("records"):
            def g(k):
                v = r.get(k, "")
                return "" if v is None or (isinstance(v, float) and pd.isna(v)) else str(v)
            lid = lead_id(event, g("Company"), g("Phone"), g("Email"))
            row = {"Lead ID": lid, "Event": event, "Scanned On": now, "Company": g("Company"),
                   "Contact Person": g("Contact Person"), "Phone": g("Phone"), "Email": g("Email"), "City": g("City"),
                   "Website": g("Website Checked") or g("Website"), "Lead Priority": g("Lead Priority"),
                   "Opportunities": g("Opportunities"), "Coverage Prospect": g("Coverage Prospect"),
                   "Pitch Note": g("Pitch Note"), "WhatsApp Link": g("WhatsApp Link"),
                   "Status": "New", "Follow-up Date": "", "Notes": "", "Updated": now}
            if lid in old_by_id:
                for c in USER_COLS:
                    row[c] = old_by_id[lid].get(c, "") or row[c]
                del old_by_id[lid]
            new_rows.append(row)
        merged = pd.DataFrame(list(old_by_id.values()) + new_rows, columns=LEAD_COLS)
        self.save(merged)
        return len(new_rows)

    def update_rows(self, edited: pd.DataFrame) -> None:
        """Save edits made in the tracker table (matched by Lead ID)."""
        all_df = self.load()
        now = f"{dt.datetime.now():%Y-%m-%d %H:%M}"
        ed = edited.set_index("Lead ID")
        for i, r in all_df.iterrows():
            lid = r["Lead ID"]
            if lid in ed.index:
                changed = False
                for c in USER_COLS:
                    v = ed.at[lid, c]
                    v = "" if v is None or (isinstance(v, float) and pd.isna(v)) else (
                        v.strftime("%Y-%m-%d") if hasattr(v, "strftime") else str(v))
                    if v == "NaT":
                        v = ""
                    if str(all_df.at[i, c]) != v:
                        all_df.at[i, c] = v
                        changed = True
                if changed:
                    all_df.at[i, "Updated"] = now
        self.save(all_df)
