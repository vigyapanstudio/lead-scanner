"""Cleans a contact list before scanning: fixes phone numbers and email typos, removes duplicates."""
from __future__ import annotations

import re

import pandas as pd

from auditor import clean_phone, is_empty

EMAIL_TYPOS = {
    "gamil.com": "gmail.com", "gmial.com": "gmail.com", "gmai.com": "gmail.com", "gmail.co": "gmail.com",
    "gmail.con": "gmail.com", "gmail.cm": "gmail.com", "gmaill.com": "gmail.com", "gnail.com": "gmail.com",
    "yahooo.com": "yahoo.com", "yaho.com": "yahoo.com", "yahoo.con": "yahoo.com", "yahoo.co": "yahoo.co.in",
    "hotmial.com": "hotmail.com", "hotmai.com": "hotmail.com", "outlok.com": "outlook.com",
    "rediffmail.con": "rediffmail.com", "redifmail.com": "rediffmail.com",
}
EMAIL_RE = re.compile(r"^[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}$")


def _fix_email(v: str) -> tuple[str, str]:
    """Returns (email, problem) where problem is '' if fine, 'fixed' if corrected, else a message."""
    if is_empty(v):
        return "", ""
    raw = str(v).strip()
    e = re.split(r"[\s,;/]+", raw.lower())[0].strip(".")
    dom = e.split("@")[-1] if "@" in e else ""
    if dom.endswith(".con"):
        e = e[:-4] + ".com"
    dom = e.split("@")[-1] if "@" in e else ""
    if dom in EMAIL_TYPOS:
        e = e.replace("@" + dom, "@" + EMAIL_TYPOS[dom])
    if not EMAIL_RE.match(e):
        return raw, "Email looks wrong"
    return e, ("fixed" if e != raw.lower().split()[0].strip(".") else "")


def clean_contacts(df: pd.DataFrame, mapping: dict) -> tuple[pd.DataFrame, pd.DataFrame, dict]:
    """Returns (clean_df, removed_duplicates_df, stats). Adds a 'Data Check' column with any warnings."""
    df = df.copy()
    for c in df.columns:
        df[c] = df[c].astype(str).str.strip().replace({"nan": ""})
    stats = {"phones_fixed": 0, "emails_fixed": 0, "bad_emails": 0, "bad_phones": 0, "duplicates": 0}
    checks = [[] for _ in range(len(df))]
    pc, ec, cc, wc = mapping.get("phone"), mapping.get("email"), mapping.get("company"), mapping.get("website")

    if pc:
        new = []
        for i, v in enumerate(df[pc]):
            p = clean_phone(v)
            if is_empty(v):
                new.append("")
            elif p:
                if p != str(v).strip():
                    stats["phones_fixed"] += 1
                new.append(p)
            else:
                stats["bad_phones"] += 1
                checks[i].append("Phone number looks wrong")
                new.append(v)
        df[pc] = new
    if ec:
        new = []
        for i, v in enumerate(df[ec]):
            e, prob = _fix_email(v)
            if prob == "fixed":
                stats["emails_fixed"] += 1
            elif prob:
                stats["bad_emails"] += 1
                checks[i].append(prob)
            new.append(e)
        df[ec] = new

    # duplicates: same phone, or same email, or same company + website
    seen, keep = set(), []
    for i, r in enumerate(df.to_dict("records")):
        keys = []
        if pc and r.get(pc) and clean_phone(r[pc]):
            keys.append("p:" + clean_phone(r[pc]))
        if ec and r.get(ec) and "@" in r[ec]:
            keys.append("e:" + r[ec].lower())
        if cc and r.get(cc):
            keys.append("c:" + re.sub(r"[^a-z0-9]", "", r[cc].lower()) + "|" +
                        (re.sub(r"^https?://(www\.)?", "", r.get(wc, "").lower()).strip("/") if wc else ""))
        dup = any(k in seen for k in keys)
        keep.append(not dup)
        seen.update(keys)
    keep_s = pd.Series(keep, index=df.index)
    removed = df[~keep_s].copy()
    stats["duplicates"] = len(removed)
    df["Data Check"] = [", ".join(c) for c in checks]
    clean = df[keep_s].reset_index(drop=True)
    return clean, removed.reset_index(drop=True), stats


def summary_text(stats: dict) -> str:
    parts = []
    if stats["duplicates"]:
        parts.append(f"removed {stats['duplicates']} duplicate contact(s)")
    if stats["phones_fixed"]:
        parts.append(f"tidied {stats['phones_fixed']} phone number(s)")
    if stats["emails_fixed"]:
        parts.append(f"fixed {stats['emails_fixed']} email typo(s)")
    if stats["bad_phones"]:
        parts.append(f"{stats['bad_phones']} phone number(s) look wrong")
    if stats["bad_emails"]:
        parts.append(f"{stats['bad_emails']} email(s) look wrong")
    return ("Cleaned your list: " + "; ".join(parts) + ".") if parts else "Your list looks clean - no duplicates found."
