"""Visiting-card reader: photo -> contact details.

Uses free Tesseract OCR by default. If an Anthropic (Claude) API key is set in Secrets, it uses
Claude's vision model instead, which reads cards much more accurately."""
from __future__ import annotations

import base64
import io
import json
import re

import requests
from PIL import Image, ImageOps

from auditor import clean_phone

CARD_FIELDS = ["Company", "Contact Person", "Designation", "Phone", "Email", "Website", "Instagram", "City"]
COMPANY_WORDS = r"\b(pvt|private|ltd|limited|llp|inc|industries|industry|enterprises?|exports?|imports?|traders?|trading|" \
                r"solutions|technologies|tech|studio|group|co\.|company|corporation|associates|agency|foods|textiles|" \
                r"fabrics|furniture|interiors|jewell?ers|steels?|engineering|works|mart|store|creations|designs|events)\b"
DESIGNATIONS = r"\b(ceo|founder|co-founder|director|manager|proprietor|owner|partner|md|head|executive|sales|marketing|" \
               r"president|chairman|consultant|designer|engineer)\b"
CITIES = ["Delhi", "New Delhi", "Mumbai", "Bengaluru", "Bangalore", "Chennai", "Kolkata", "Hyderabad", "Pune", "Ahmedabad",
          "Jaipur", "Ludhiana", "Chandigarh", "Noida", "Gurugram", "Gurgaon", "Surat", "Indore", "Lucknow", "Kanpur",
          "Amritsar", "Jalandhar", "Jodhpur", "Rajkot", "Vadodara", "Nagpur", "Bhopal", "Coimbatore", "Kochi", "Agra",
          "Moradabad", "Panipat", "Faridabad", "Ghaziabad", "Mohali", "Dehradun", "Patna", "Guwahati", "Goa"]


def _prep(img_bytes: bytes) -> Image.Image:
    img = Image.open(io.BytesIO(img_bytes))
    img = ImageOps.exif_transpose(img).convert("L")
    if max(img.size) < 1800:
        f = 1800 / max(img.size)
        img = img.resize((int(img.width * f), int(img.height * f)))
    return ImageOps.autocontrast(img)


def ocr_text(img_bytes: bytes) -> str:
    import pytesseract
    return pytesseract.image_to_string(_prep(img_bytes), config="--psm 6")


def parse_card_text(text: str) -> dict:
    lines = [l.strip(" |:-•*") for l in text.splitlines() if l.strip(" |:-•*")]
    joined = " ".join(lines)
    out = {k: "" for k in CARD_FIELDS}
    m = re.search(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}", joined)
    if m:
        out["Email"] = m.group(0).lower()
    m = re.search(r"(?:https?://)?(?:www\.)[A-Za-z0-9.-]+\.[A-Za-z]{2,}(?:/\S*)?", joined, re.I) or \
        re.search(r"\b[A-Za-z0-9-]+\.(?:com|in|co\.in|net|org|biz|store|shop)\b(?!@)", joined.replace(out["Email"], ""), re.I)
    if m:
        out["Website"] = m.group(0).lower()
    m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)|(?:^|\s)@([A-Za-z0-9_.]{3,30})\b", joined, re.I)
    if m:
        out["Instagram"] = "@" + (m.group(1) or m.group(2))
    phones = re.findall(r"(?:\+?91[\s-]?)?[6-9]\d{2,4}[\s-]?\d{2,4}[\s-]?\d{2,4}", joined)
    phones = [clean_phone(p) for p in phones if clean_phone(p)]
    out["Phone"] = ", ".join(dict.fromkeys(phones))
    for c in CITIES:
        if re.search(rf"\b{c}\b", joined, re.I):
            out["City"] = c
            break
    # company: a line with business words, else the biggest ALL-CAPS line
    text_lines = [l for l in lines if not re.search(r"@|www\.|\.com|\d{5,}", l, re.I)]
    comp = [l for l in text_lines if re.search(COMPANY_WORDS, l, re.I)]
    if comp:
        out["Company"] = comp[0]
    else:
        caps = [l for l in text_lines if l.isupper() and len(l) > 3]
        if caps:
            out["Company"] = max(caps, key=len).title()
    # designation + name: a short line of 2-3 words with no digits, near a designation line
    for i, l in enumerate(text_lines):
        if re.search(DESIGNATIONS, l, re.I) and not out["Designation"]:
            out["Designation"] = l
            if i > 0 and not re.search(COMPANY_WORDS, text_lines[i - 1], re.I):
                out["Contact Person"] = text_lines[i - 1]
    if not out["Contact Person"]:
        for l in text_lines:
            w = l.split()
            if 2 <= len(w) <= 3 and all(x[:1].isupper() for x in w) and not re.search(COMPANY_WORDS + "|" + DESIGNATIONS, l, re.I) \
                    and l != out["Company"]:
                out["Contact Person"] = l
                break
    if not out["Company"] and out["Website"]:
        out["Company"] = re.sub(r"^(https?://)?(www\.)?", "", out["Website"]).split(".")[0].title()
    return out


def claude_read_card(img_bytes: bytes, api_key: str, model: str) -> dict | None:
    img = _prep(img_bytes).convert("RGB")
    img.thumbnail((1600, 1600))
    b = io.BytesIO()
    img.save(b, "JPEG", quality=88)
    prompt = ("This is a photo of a business/visiting card. Return ONLY a JSON object with these keys: "
              + ", ".join(f'"{k}"' for k in CARD_FIELDS)
              + '. Use "" for anything not on the card. Phone: digits only, comma-separate if several.')
    try:
        r = requests.post("https://api.anthropic.com/v1/messages", timeout=60,
                          headers={"x-api-key": api_key, "anthropic-version": "2023-06-01",
                                   "content-type": "application/json"},
                          json={"model": model, "max_tokens": 500, "messages": [{"role": "user", "content": [
                              {"type": "image", "source": {"type": "base64", "media_type": "image/jpeg",
                                                           "data": base64.b64encode(b.getvalue()).decode()}},
                              {"type": "text", "text": prompt}]}]})
        if r.status_code != 200:
            return None
        txt = "".join(c.get("text", "") for c in r.json().get("content", []))
        data = json.loads(re.search(r"\{.*\}", txt, re.S).group(0))
        return {k: str(data.get(k, "") or "") for k in CARD_FIELDS}
    except Exception:
        return None


def read_card(img_bytes: bytes, api_key: str = "", model: str = "") -> tuple[dict, str]:
    """Returns (fields, method_used)."""
    if api_key and model:
        res = claude_read_card(img_bytes, api_key, model)
        if res:
            return res, "AI"
    try:
        return parse_card_text(ocr_text(img_bytes)), "OCR"
    except Exception:
        return {k: "" for k in CARD_FIELDS}, "failed"
