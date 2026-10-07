"""A 1-page 'Free Digital Health Check' PDF for each company - to send to the lead itself."""
from __future__ import annotations

import datetime as dt
import io
import re
import zipfile
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.platypus import Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle

from brand import CLIENT_EMAIL, CLIENT_INSTAGRAM, CLIENT_NAME, CLIENT_PHONE, CLIENT_WEBSITE
from pdf_report import (ACCENT, ACCENT_BG, CONTENT_W, GREY, GRN, LINE, MARGIN, NAVY, PANEL, RED, S, esc,
                        score_bar, score_colors, val)

ST = {
    "brand": S("m_brand", fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=ACCENT),
    "tag": S("m_tag", fontSize=8, leading=10, textColor=colors.HexColor("#9DBFB9")),
    "h1": S("m_h1", fontName="Helvetica-Bold", fontSize=21, leading=25, textColor=colors.white),
    "sub": S("m_sub", fontSize=9, leading=12, textColor=colors.HexColor("#C9D8D4")),
    "h2": S("m_h2", fontName="Helvetica-Bold", fontSize=12.5, leading=16, spaceBefore=4, spaceAfter=4),
    "body": S("m_body", fontSize=9.5, leading=13.5),
    "bullet": S("m_bullet", fontSize=9.5, leading=13, leftIndent=10, firstLineIndent=-10),
    "label": S("m_label", fontSize=7.5, leading=9.5, textColor=GREY),
    "score": S("m_score", fontName="Helvetica-Bold", fontSize=22, leading=26),
    "small": S("m_small", fontSize=8.5, leading=11),
    "cta": S("m_cta", fontName="Helvetica-Bold", fontSize=12, leading=16),
}

WHY = {
    "New Website": "Most customers look a business up online before they call. Without a working website, they move on to a competitor.",
    "Website Redesign": "An old-looking or non-mobile website makes people doubt the business - and most visitors are on their phones.",
    "Website Upgrade": "A few fixes (security, speed, mobile layout) would make the website feel more trustworthy.",
    "SEO": "If Google can't read the website properly, customers searching for what you sell will find others first.",
    "Instagram Page Setup": "Instagram is where many buyers discover and check brands today. No page means missed attention.",
    "Instagram Management": "An inactive Instagram can make the business look closed or uninterested.",
    "Google Business Profile": "A strong Google Maps listing with reviews brings in local customers searching nearby.",
    "Social Media Setup": "Being present on Facebook, LinkedIn or YouTube builds trust with different kinds of buyers.",
}
HELP = {
    "New Website": "A modern, mobile-friendly website that brings enquiries",
    "Website Redesign": "A fresh redesign that looks great on phones",
    "Website Upgrade": "Quick fixes: security (HTTPS), speed and mobile layout",
    "SEO": "Google setup so customers can find you (SEO)",
    "Instagram Page Setup": "Instagram page setup with a content plan",
    "Instagram Management": "Regular Instagram posts and reels for your brand",
    "Google Business Profile": "Google Maps listing setup and review growth",
    "Social Media Setup": "Facebook, LinkedIn and YouTube presence",
}


def _P(t, st):
    return Paragraph(t, ST[st])


def _issues(row) -> list[str]:
    out = []
    for col, limit in (("Website Issues", 4), ("SEO Issues", 3), ("Instagram Issues", 1), ("Social Media Issues", 2)):
        items = [i.strip().lstrip("•").strip() for i in val(row.get(col)).split("\n") if i.strip()]
        out += items[:limit]
    # drop scanner-only wording that isn't useful to the business owner
    out = [i for i in out if not i.lower().startswith(("scanner error", "website blocked our automatic"))]
    return out[:9]


def _contact_lines() -> str:
    bits = [b for b in (CLIENT_PHONE, CLIENT_EMAIL, CLIENT_WEBSITE, CLIENT_INSTAGRAM) if b]
    return " &nbsp;·&nbsp; ".join(escape(b) for b in bits) if bits else "Just reply to our message and we'll take it from there."


def build_company_pdf(row: dict) -> bytes:
    company = esc(row.get("Company") or row.get("Contact Person"), "Your business")
    person = val(row.get("Contact Person"))
    today = f"{dt.date.today():%d %B %Y}"
    opp = [o.strip() for o in val(row.get("Opportunities")).split(",") if o.strip() and "None" not in o
           and o.strip() != "Check manually"]
    healthy = not opp

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN, topMargin=12 * mm,
                            bottomMargin=12 * mm, title=f"Digital Health Check - {val(row.get('Company'))}",
                            author=CLIENT_NAME)
    story = []

    # header band
    head = Table([[_P(escape(CLIENT_NAME.upper()), "brand")],
                  [_P("FREE DIGITAL HEALTH CHECK", "tag")],
                  [_P(f"Prepared for {company}", "h1")],
                  [_P((f"Attention: {escape(person)} &nbsp;·&nbsp; " if person else "") + today, "sub")]],
                 colWidths=[CONTENT_W])
    head.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("LEFTPADDING", (0, 0), (-1, -1), 14),
                              ("TOPPADDING", (0, 0), (-1, 0), 14), ("BOTTOMPADDING", (0, -1), (-1, -1), 14),
                              ("TOPPADDING", (0, 1), (-1, -1), 2)]))
    story += [head, Spacer(1, 6 * mm)]

    intro = ("Good news - your business is in strong shape online. Here's a quick summary of what we checked."
             if healthy else
             "We ran a quick, free check of how your business appears online - the way a new customer would see it. "
             "Here's what we found, in plain language.")
    story += [_P(intro, "body"), Spacer(1, 5 * mm)]

    # scores
    def box(title, score, note):
        f2, b2 = score_colors(score)
        sc = val(score)
        parts = [_P(escape(title).upper(), "label")]
        if sc:
            parts += [_P(f'<font color="{f2.hexval()}">{sc}</font><font size="10" color="{GREY.hexval()}"> /100</font>', "score"),
                      score_bar(sc, CONTENT_W / 3 - 16 * mm, 5), Spacer(1, 3)]
        else:
            parts += [_P(f'<font size="13" color="{RED.hexval()}">{escape(note)}</font>', "score")]
            note = ""
        if note:
            parts.append(_P(escape(note), "small"))
        return parts, (b2 if sc else colors.HexColor("#FDE8E6"))

    ws = val(row.get("Website Status"))
    web_note = {"No Website": "No website found", "Broken / Not Opening": "Website not opening",
                "Parked / Empty / Error": "Website is empty"}.get(ws, ws)
    w_box, w_bg = box("Website", row.get("Website Score"), web_note)
    seo_sc = row.get("SEO Score") if ws not in ("No Website", "Broken / Not Opening") else None
    seo_note = {"Needs SEO Work": "Hard to find on Google", "Basic SEO OK": "Basics in place",
                "Good": "Easy for Google to read"}.get(val(row.get("SEO Status")), "Not visible")
    s_box, s_bg = box("Google visibility", seo_sc, seo_note if val(seo_sc) else "Not visible")
    have = [k for k in ("Facebook", "LinkedIn", "YouTube") if val(row.get(k))]
    ig = val(row.get("Instagram Handle"))
    social_line = ("Instagram: " + ("@" + ig if ig else "not found")) + (f" · also on {', '.join(have)}" if have else "")
    soc_ok = bool(ig)
    soc = [_P("SOCIAL MEDIA", "label"),
           _P(f'<font size="13" color="{(GRN if soc_ok else RED).hexval()}">{"Present" if soc_ok else "Missing"}</font>', "score"),
           _P(escape(social_line), "small")]
    scores = Table([[w_box, s_box, soc]], colWidths=[CONTENT_W / 3] * 3)
    scores.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("BACKGROUND", (0, 0), (0, 0), w_bg),
                                ("BACKGROUND", (1, 0), (1, 0), s_bg),
                                ("BACKGROUND", (2, 0), (2, 0), colors.HexColor("#E2F3E7") if soc_ok else colors.HexColor("#FDE8E6")),
                                ("LINEAFTER", (0, 0), (1, 0), 4, colors.white), ("LEFTPADDING", (0, 0), (-1, -1), 10),
                                ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    story += [scores, Spacer(1, 6 * mm)]

    # findings
    issues = _issues(row)
    story.append(_P("What we found", "h2"))
    if issues:
        story += [_P(f"• {esc(i)}", "bullet") for i in issues]
    else:
        story.append(_P("No major problems - your website and social media cover the basics well.", "body"))
    story.append(Spacer(1, 5 * mm))

    if healthy:
        story += [_P("An idea for you", "h2"),
                  _P(f"Brands that are already strong online get even more out of being seen at events. "
                     f"{escape(CLIENT_NAME)} covers exhibitions and events across India through video coverage and "
                     f"podcasts - we'd love to feature {company}.", "body")]
    else:
        story.append(_P("Why it matters", "h2"))
        story += [_P(f"• {escape(WHY[o])}", "bullet") for o in opp if o in WHY][:3]
        story.append(Spacer(1, 5 * mm))
        story.append(_P("How we can help", "h2"))
        story += [_P(f"• <b>{escape(HELP[o])}</b>", "bullet") for o in opp if o in HELP]
    story.append(Spacer(1, 7 * mm))

    cta = Table([[_P("Want to talk about it?" if not healthy else "Interested in a feature?", "cta")],
                 [_P(f"<b>{escape(CLIENT_NAME)}</b> &nbsp;·&nbsp; " + _contact_lines(), "small")]],
                colWidths=[CONTENT_W])
    cta.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), ACCENT_BG), ("BOX", (0, 0), (-1, -1), 1, ACCENT),
                             ("LEFTPADDING", (0, 0), (-1, -1), 14), ("TOPPADDING", (0, 0), (-1, 0), 12),
                             ("BOTTOMPADDING", (0, -1), (-1, -1), 12)]))
    story += [cta, Spacer(1, 5 * mm),
              _P(f'<font color="{GREY.hexval()}">This free check was done automatically on {today} using only public '
                 'information from your home page and social media links. It is a first look, not a full audit.</font>',
                 "small")]
    doc.build(story)
    return buf.getvalue()


def safe_name(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9 &()._-]+", "", s or "").strip()
    return (s or "Company")[:60]


def build_zip(results: pd.DataFrame, priorities=("HOT", "WARM", "LOW", "CHECK")) -> tuple[bytes, int]:
    buf = io.BytesIO()
    used = set()
    count = 0
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for r in results.to_dict("records"):
            if r.get("Lead Priority") not in priorities:
                continue
            name = safe_name(val(r.get("Company")) or val(r.get("Contact Person")))
            base, k = name, 2
            while name.lower() in used:
                name = f"{base} ({k})"
                k += 1
            used.add(name.lower())
            z.writestr(f"{name} - Digital Health Check.pdf", build_company_pdf(r))
            count += 1
    return buf.getvalue(), count
