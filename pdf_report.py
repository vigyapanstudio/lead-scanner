"""Builds a clear, printable PDF report: summary + contact index + one card per contact."""
from __future__ import annotations

import datetime as dt
import io
from xml.sax.saxutils import escape

import pandas as pd
from reportlab.graphics.shapes import Drawing, Rect, String
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from brand import CLIENT_NAME, POWERED_BY
from reportlab.platypus import (KeepTogether, PageBreak, Paragraph, SimpleDocTemplate, Spacer,
                                Table, TableStyle)

NAVY = colors.HexColor("#0B2E2C")
ACCENT = colors.HexColor("#FFB72B")
ACCENT_BG = colors.HexColor("#FFF4DB")
GREY = colors.HexColor("#6B7280")
LINE = colors.HexColor("#E2E5EC")
PANEL = colors.HexColor("#F6F7FA")
RED, RED_BG = colors.HexColor("#B42318"), colors.HexColor("#FDE8E6")
AMB, AMB_BG = colors.HexColor("#8A5A00"), colors.HexColor("#FFF3D1")
GRN, GRN_BG = colors.HexColor("#1E6B37"), colors.HexColor("#E2F3E7")
PRIO = {"HOT": (RED, RED_BG), "WARM": (AMB, AMB_BG), "LOW": (GRN, GRN_BG)}

PAGE_W, PAGE_H = A4
MARGIN = 14 * mm
CONTENT_W = PAGE_W - 2 * MARGIN


def S(name, **kw):
    base = dict(fontName="Helvetica", fontSize=9, leading=12, textColor=NAVY)
    base.update(kw)
    return ParagraphStyle(name, **base)


ST = {
    "title": S("title", fontName="Helvetica-Bold", fontSize=24, leading=29),
    "sub": S("sub", fontSize=10, textColor=GREY, leading=14),
    "h2": S("h2", fontName="Helvetica-Bold", fontSize=14, leading=18, spaceBefore=6, spaceAfter=6),
    "body": S("body"),
    "small": S("small", fontSize=8, leading=10.5),
    "tiny": S("tiny", fontSize=7.2, leading=9),
    "tinyb": S("tinyb", fontName="Helvetica-Bold", fontSize=7.2, leading=9),
    "label": S("label", fontSize=7, leading=9, textColor=GREY),
    "value": S("value", fontName="Helvetica-Bold", fontSize=9, leading=11.5),
    "cardhead": S("cardhead", fontName="Helvetica-Bold", fontSize=12, leading=15, textColor=colors.white),
    "cardsub": S("cardsub", fontSize=8, leading=10, textColor=colors.HexColor("#C9CFDB")),
    "sec": S("sec", fontName="Helvetica-Bold", fontSize=8.5, leading=11, textColor=NAVY, spaceAfter=2),
    "bullet": S("bullet", fontSize=8, leading=10.5, leftIndent=8, firstLineIndent=-8),
    "pitch": S("pitch", fontSize=9, leading=12.5),
    "kpi": S("kpi", fontName="Helvetica-Bold", fontSize=22, leading=26, alignment=TA_CENTER),
    "kpil": S("kpil", fontSize=8, leading=10, textColor=GREY, alignment=TA_CENTER),
    "badge": S("badge", fontName="Helvetica-Bold", fontSize=8.5, leading=11, alignment=TA_CENTER),
    "score": S("score", fontName="Helvetica-Bold", fontSize=18, leading=21),
    "right": S("right", fontSize=8, leading=10, textColor=GREY, alignment=TA_RIGHT),
}

_SAFE = str.maketrans({"→": "->", "–": "-", "—": "-", "‘": "'", "’": "'",
                       "“": '"', "”": '"', "…": "...", "₹": "Rs."})


def val(v, default="") -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    s = str(v).strip()
    if s.lower() in ("nan", "none", "nat"):
        return default
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s


def esc(v, default="-") -> str:
    s = val(v, default).translate(_SAFE)
    # keep only characters the built-in PDF font can draw
    s = "".join(ch if ord(ch) < 256 or ch == "•" else "?" for ch in s)
    return escape(s)


def P(text, style="body"):
    return Paragraph(text, ST[style] if isinstance(style, str) else style)


def score_colors(score):
    try:
        s = float(score)
    except (TypeError, ValueError):
        return GREY, PANEL
    if pd.isna(s):
        return GREY, PANEL
    if s < 50:
        return RED, RED_BG
    if s < 80:
        return AMB, AMB_BG
    return GRN, GRN_BG


def score_bar(score, width=60 * mm, height=4):
    d = Drawing(width, height + 1)
    d.add(Rect(0, 0, width, height, fillColor=LINE, strokeColor=None))
    try:
        s = max(0, min(100, float(score)))
        fg, _ = score_colors(s)
        d.add(Rect(0, 0, width * s / 100, height, fillColor=fg, strokeColor=None))
    except (TypeError, ValueError):
        pass
    return d


def badge(text, fg, bg, width=22 * mm):
    t = Table([[P(f'<font color="{fg.hexval()}">{escape(text)}</font>', "badge")]],
              colWidths=[width], rowHeights=[16])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), bg),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("ROUNDEDCORNERS", [4, 4, 4, 4]),
                           ("TOPPADDING", (0, 0), (-1, -1), 2), ("BOTTOMPADDING", (0, 0), (-1, -1), 2)]))
    return t


def issues_to_bullets(text, empty_msg):
    s = val(text)
    items = [i.strip().lstrip("•").strip() for i in s.split("\n") if i.strip()]
    if not items:
        return [P(f'<font color="{GRN.hexval()}">{escape(empty_msg)}</font>', "bullet")]
    return [P(f"• {esc(i)}", "bullet") for i in items]


# ----------------------------------------------------------------------------- page decoration
def _decorate(source, stamp):
    def draw(canvas, doc):
        canvas.saveState()
        canvas.setFont("Helvetica-Bold", 8)
        canvas.setFillColor(NAVY)
        canvas.drawString(MARGIN, PAGE_H - 9 * mm, f"{CLIENT_NAME.upper()}  \u00b7  Lead Scanner Report")
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(GREY)
        canvas.drawRightString(PAGE_W - MARGIN, PAGE_H - 9 * mm, f"Page {doc.page}")
        canvas.setStrokeColor(ACCENT)
        canvas.setLineWidth(1.2)
        canvas.line(MARGIN, PAGE_H - 10.5 * mm, PAGE_W - MARGIN, PAGE_H - 10.5 * mm)
        src = source if len(source) <= 60 else source[:57] + "..."
        canvas.drawString(MARGIN, 7 * mm, f"Source: {src}   ·   Scanned {stamp}")
        canvas.drawRightString(PAGE_W - MARGIN, 7 * mm, f"Powered by {POWERED_BY}")
        canvas.restoreState()
    return draw


# ----------------------------------------------------------------------------- sections
def _summary(df, source, stamp):
    out = [P(f'<font color="#B57A00"><b>{escape(CLIENT_NAME.upper())}</b></font>', "label"),
           P("Lead Scanner Report", "title"),
           P(f"{escape(source)}  ·  {len(df)} contacts scanned  ·  {stamp}", "sub"),
           Spacer(1, 8 * mm)]
    opp = df["Opportunities"].fillna("")
    tiles = [
        ("Contacts scanned", len(df), NAVY),
        ("HOT leads", int((df["Lead Priority"] == "HOT").sum()), RED),
        ("WARM leads", int((df["Lead Priority"] == "WARM").sum()), AMB),
        ("Healthy (LOW)", int((df["Lead Priority"] == "LOW").sum()), GRN),
        ("No website", int(opp.str.contains("New Website").sum()), NAVY),
        ("Website redesign / upgrade", int(opp.str.contains("Redesign|Upgrade").sum()), NAVY),
        ("Need SEO work", int(opp.str.contains("SEO").sum()), NAVY),
        ("Instagram leads", int(opp.str.contains("Instagram").sum()), NAVY),
    ]
    cells = [[P(f'<font color="{c.hexval()}">{n}</font>', "kpi"), P(escape(l), "kpil")] for l, n, c in tiles]
    w = CONTENT_W / 4
    grid = Table([[cells[i] for i in range(4)], [cells[i] for i in range(4, 8)]],
                 colWidths=[w] * 4, rowHeights=[22 * mm, 22 * mm])
    grid.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), PANEL), ("BOX", (0, 0), (-1, -1), 0.6, LINE),
        ("INNERGRID", (0, 0), (-1, -1), 3, colors.white), ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
    out += [grid, Spacer(1, 9 * mm)]

    # bar chart of services
    services = [("New website", "New Website"), ("Website redesign", "Website Redesign"),
                ("Website upgrade", "Website Upgrade"), ("SEO", "SEO"),
                ("Instagram page setup", "Instagram Page Setup"),
                ("Instagram management", "Instagram Management")]
    counts = [(l, int(opp.str.contains(k).sum())) for l, k in services]
    mx = max([c for _, c in counts] + [1])
    row_h, label_w = 9 * mm, 45 * mm
    bar_w = CONTENT_W - label_w - 20 * mm
    d = Drawing(CONTENT_W, row_h * len(counts))
    for i, (l, c) in enumerate(counts):
        y = row_h * (len(counts) - 1 - i) + 2.5 * mm
        d.add(String(0, y + 1, l, fontName="Helvetica", fontSize=9, fillColor=NAVY))
        d.add(Rect(label_w, y - 1, bar_w, 4.5 * mm, fillColor=PANEL, strokeColor=None))
        d.add(Rect(label_w, y - 1, max(1, bar_w * c / mx), 4.5 * mm, fillColor=ACCENT, strokeColor=None))
        pct = f"  ({round(100 * c / len(df))}%)" if len(df) else ""
        d.add(String(label_w + bar_w + 3, y + 1, f"{c}{pct}", fontName="Helvetica-Bold",
                     fontSize=9, fillColor=NAVY))
    out += [P("Services these leads need", "h2"), d, Spacer(1, 8 * mm)]

    # how to read
    legend = [
        [badge("HOT", RED, RED_BG), P("Needs 2 or more big services (for example no website <b>and</b> no Instagram). Call these first.", "small")],
        [badge("WARM", AMB, AMB_BG), P("Needs 1-2 services.", "small")],
        [badge("LOW", GRN, GRN_BG), P("Already in good digital shape - pitch event coverage or a podcast feature instead.", "small")],
        [badge("CHECK", GREY, PANEL), P("The website blocked our automatic visit - open it yourself to judge it.", "small")],
        [P("<b>Scores</b>", "small"), P(f'Website and SEO are scored out of 100: <font color="{RED.hexval()}"><b>red under 50</b></font> = poor, '
                                        f'<font color="{AMB.hexval()}"><b>yellow 50-79</b></font> = average, '
                                        f'<font color="{GRN.hexval()}"><b>green 80+</b></font> = good.', "small")],
        [P("<b>Website age</b>", "small"), P("Latest year seen in the website footer (©) or server date. 'Old' = no sign of an update in 4+ years.", "small")],
        [P("<b>Please note</b>", "small"), P("This is an automated first check of each company's home page. Open the website yourself before pitching. 'Check Manually' means the website blocked our automatic visit.", "small")],
        [P("<b>Instagram</b>", "small"), P("Found from your list or from a link on their website. Followers / last post are shown only when the detailed Instagram check is switched on.", "small")],
    ]
    lt = Table(legend, colWidths=[26 * mm, CONTENT_W - 26 * mm])
    lt.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                            ("BOTTOMPADDING", (0, 0), (-1, -1), 5), ("TOPPADDING", (0, 0), (-1, -1), 5),
                            ("LINEBELOW", (0, 0), (-1, -2), 0.4, LINE)]))
    out += [P("How to read this report", "h2"), lt]
    return out


def _index(df):
    head = ["#", "Company / Contact", "Phone", "Priority", "Website", "Web", "SEO", "Instagram", "Needs"]
    rows = [[P(f'<font color="#FFFFFF"><b>{h}</b></font>', "tinyb") for h in head]]
    style = [("BACKGROUND", (0, 0), (-1, 0), NAVY), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
             ("VALIGN", (0, 0), (-1, -1), "MIDDLE"), ("LINEBELOW", (0, 0), (-1, -1), 0.4, LINE),
             ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
             ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]
    for i, r in enumerate(df.itertuples(index=False), 1):
        r = r._asdict()
        name = esc(r.get("Company") or r.get("Contact_Person"))
        person = esc(r.get("Contact_Person"), "")
        sub = f'<br/><font color="{GREY.hexval()}">{person}</font>' if person and person != name else ""
        pr = val(r.get("Lead_Priority"))
        web, seo = val(r.get("Website_Score")), val(r.get("SEO_Score"))
        needs = val(r.get("Opportunities")).replace("None - digitally healthy", "-")
        rows.append([P(str(i), "tiny"), P(f"<b>{name}</b>{sub}", "tiny"), P(esc(r.get("Phone")), "tiny"),
                     P(f'<font color="{PRIO.get(pr, (GREY, PANEL))[0].hexval()}"><b>{escape(pr)}</b></font>', "tinyb"),
                     P(esc(r.get("Website_Status")), "tiny"),
                     P(f"<b>{web or '-'}</b>", "tinyb"), P(f"<b>{seo or '-'}</b>", "tinyb"),
                     P(esc(r.get("Instagram_Status")), "tiny"), P(esc(needs), "tiny")])
        fg, bg = PRIO.get(pr, (GREY, PANEL))
        style += [("BACKGROUND", (3, i), (3, i), bg), ("TEXTCOLOR", (3, i), (3, i), fg)]
        for col, sc in ((5, web), (6, seo)):
            if sc:
                f2, b2 = score_colors(sc)
                style.append(("BACKGROUND", (col, i), (col, i), b2))
        if i % 2 == 0:
            style.append(("BACKGROUND", (0, i), (2, i), PANEL))
    widths = [8, 42, 22, 15, 27, 10, 10, 25, 23]
    k = CONTENT_W / (sum(widths) * mm)
    t = Table(rows, colWidths=[w * mm * k for w in widths], repeatRows=1)
    t.setStyle(TableStyle(style))
    # colour the text of priority/score cells via paragraph colour
    return [P("All contacts at a glance", "h2"),
            P("Sorted with HOT leads first. The number (#) matches the detailed card for each contact in the next section.", "sub"),
            Spacer(1, 3 * mm), t]


def _card(i, r, extras):
    pr = val(r.get("Lead Priority"))
    fg, bg = PRIO.get(pr, (GREY, PANEL))
    company = esc(r.get("Company") or r.get("Contact Person"), "Unnamed contact")
    excel_row = val(r.get("_excel_row"))

    # header
    head = Table([[P(f"#{i}  {company}", "cardhead"),
                   badge(f"{pr} LEAD" if pr else "-", fg, bg, 26 * mm)],
                  [P(f"Row {excel_row} in your Excel file", "cardsub"), ""]],
                 colWidths=[CONTENT_W - 32 * mm, 32 * mm])
    head.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), NAVY), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                              ("SPAN", (1, 0), (1, 1)), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                              ("TOPPADDING", (0, 0), (-1, 0), 7), ("BOTTOMPADDING", (0, -1), (-1, -1), 7)]))

    # contact details
    fields = [("Contact person", r.get("Contact Person")), ("Phone", r.get("Phone")),
              ("Email", r.get("Email")), ("Website", r.get("Website Checked") or r.get("Website")),
              ("City", r.get("City")), ("Industry", r.get("Category"))]
    fields += [(k, r.get(k_full)) for k, k_full in extras]
    fields = [(k, v) for k, v in fields if val(v) or k in ("Contact person", "Phone", "Email", "Website")]
    cells = [[P(escape(k), "label"),
              P(esc(v) if val(v) else f'<font color="{GREY.hexval()}">Not given</font>', "value")]
             for k, v in fields]
    while len(cells) % 3:
        cells.append(["", ""])
    cw = CONTENT_W / 3
    det_rows = []
    for j in range(0, len(cells), 3):
        det_rows.append([cells[j][0], cells[j + 1][0], cells[j + 2][0]])
        det_rows.append([cells[j][1], cells[j + 1][1], cells[j + 2][1]])
    details = Table(det_rows, colWidths=[cw] * 3)
    details.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                                 ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]
                                + [("BOTTOMPADDING", (0, k), (-1, k), 5) for k in range(1, len(det_rows), 2)]))

    # three score boxes
    def box(title, score, line1, line2=""):
        f2, b2 = score_colors(score)
        sc = val(score)
        big = f'<font color="{f2.hexval()}">{sc}</font><font size="9" color="{GREY.hexval()}"> /100</font>' if sc else \
            f'<font size="11" color="{GREY.hexval()}">{esc(line1 if line1 not in ("No Website", "Could not check") else "Not checked")}</font>'
        parts = [P(escape(title).upper(), "label"), P(big, "score")]
        if sc:
            parts += [score_bar(sc, CONTENT_W / 3 - 14 * mm), Spacer(1, 2), P(f"<b>{esc(line1)}</b>", "small")]
        if line2:
            parts.append(P(esc(line2, ""), "small"))
        return parts

    ig_h = val(r.get("Instagram Handle"))
    ig_extra = []
    if val(r.get("Instagram Followers")):
        ig_extra.append(f"{val(r.get('Instagram Followers'))} followers")
    if val(r.get("Instagram Last Post")):
        ig_extra.append(f"last post {val(r.get('Instagram Last Post'))}")
    ig_box = [P("INSTAGRAM", "label"),
              P(f'<font size="11">{"@" + esc(ig_h) if ig_h else "Not found"}</font>', "score"),
              P(f"<b>{esc(r.get('Instagram Status'))}</b>", "small")]
    if ig_extra:
        ig_box.append(P(escape(", ".join(ig_extra)), "small"))
    web_line2 = val(r.get("Website Age"))
    scores = Table([[box("Website", r.get("Website Score"), val(r.get("Website Status")), web_line2),
                     box("Google / SEO", r.get("SEO Score"), val(r.get("SEO Status"))),
                     ig_box]], colWidths=[CONTENT_W / 3] * 3)
    ws_fg, ws_bg = score_colors(r.get("Website Score"))
    seo_fg, seo_bg = score_colors(r.get("SEO Score"))
    scores.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"),
                                ("BACKGROUND", (0, 0), (0, 0), ws_bg if val(r.get("Website Score")) else PANEL),
                                ("BACKGROUND", (1, 0), (1, 0), seo_bg if val(r.get("SEO Score")) else PANEL),
                                ("BACKGROUND", (2, 0), (2, 0), PANEL),
                                ("LINEAFTER", (0, 0), (1, 0), 3, colors.white),
                                ("LEFTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 6),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 7)]))

    # findings
    left = [P("What we found on the website", "sec")] + issues_to_bullets(r.get("Website Issues"), "No problems found")
    left += [Spacer(1, 4), P("Instagram", "sec")] + issues_to_bullets(r.get("Instagram Issues"), "Instagram page found")
    if val(r.get("Website Status")) in ("No Website", "Broken / Not Opening"):
        right = [P("Google / SEO", "sec"), P("Not checked - no working website.", "bullet")]
    else:
        right = [P("Google / SEO problems", "sec")] + issues_to_bullets(r.get("SEO Issues"), "SEO basics are in place")
    findings = Table([[left, right]], colWidths=[CONTENT_W / 2] * 2)
    findings.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                                  ("TOPPADDING", (0, 0), (-1, -1), 6)]))

    # services + pitch
    opp = [o.strip() for o in val(r.get("Opportunities")).split(",") if o.strip()]
    chips = " &nbsp; ".join(f'<font backColor="{ACCENT_BG.hexval()}" color="#7A5200"><b>&nbsp;{escape(o)}&nbsp;</b></font>'
                            for o in opp) if opp else "-"
    pitch = Table([[P("RECOMMENDED SERVICES", "label")], [P(chips, "small")],
                   [P("PITCH LINE", "label")], [P(esc(r.get("Pitch Note")), "pitch")]],
                  colWidths=[CONTENT_W])
    pitch.setStyle(TableStyle([("BACKGROUND", (0, 2), (-1, 3), ACCENT_BG), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                               ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 1), (-1, 1), 6),
                               ("BOTTOMPADDING", (0, 3), (-1, 3), 8), ("LINEABOVE", (0, 0), (-1, 0), 0.5, LINE)]))

    card = Table([[head], [Spacer(1, 3)], [details], [scores], [findings], [pitch]], colWidths=[CONTENT_W])
    card.setStyle(TableStyle([("BOX", (0, 0), (-1, -1), 0.8, LINE), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                              ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                              ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return KeepTogether([card, Spacer(1, 7 * mm)])


def build_pdf(results: pd.DataFrame, source_name: str = "") -> bytes:
    df = results.copy().reset_index(drop=True)
    df["_excel_row"] = range(2, len(df) + 2)          # row number in the client's original sheet
    order = {"HOT": 0, "WARM": 1, "CHECK": 2, "LOW": 3}
    df["_o"] = df["Lead Priority"].map(order).fillna(3)
    df["_s"] = pd.to_numeric(df["Website Score"], errors="coerce").fillna(-1)
    df = df.sort_values(["_o", "_s"], kind="stable").drop(columns=["_o", "_s"]).reset_index(drop=True)
    extras = [(c.replace("[Original] ", ""), c) for c in df.columns if c.startswith("[Original]")
              and not any(w in c.lower() for w in ("s.no", "sno", "sr", "remark", "serial"))][:3]

    stamp = f"{dt.datetime.now():%d %b %Y, %I:%M %p}"
    source = source_name or "contact list"
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=16 * mm, bottomMargin=13 * mm,
                            title="Lead Scanner Report", author=CLIENT_NAME)
    story = _summary(df, source, stamp) + [PageBreak()]
    # index uses attribute-style names
    idx = df.rename(columns=lambda c: c.replace(" ", "_"))
    story += _index(idx) + [PageBreak(), P("Contact details", "h2"),
                            P("One card per contact: their details, scores, every problem we found, the services to offer and a ready pitch line.", "sub"),
                            Spacer(1, 4 * mm)]
    for i, r in enumerate(df.to_dict("records"), 1):
        story.append(_card(i, r, extras))
    deco = _decorate(source, stamp)
    doc.build(story, onFirstPage=deco, onLaterPages=deco)
    return buf.getvalue()
