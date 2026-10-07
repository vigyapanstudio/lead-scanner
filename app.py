"""Eventra Vlogs - Lead Scanner website (built by Vigyapan Studio).
Upload Excel -> scan websites, SEO and Instagram -> download PDF / Excel report."""
import datetime as dt
import hmac
import os
from html import escape

import pandas as pd
import streamlit as st

import auditor
from brand import CLIENT_NAME, POWERED_BY
from pdf_report import build_pdf
from report import build_report

st.set_page_config(page_title=f"{CLIENT_NAME} · Lead Scanner", page_icon="▶️", layout="wide",
                   initial_sidebar_state="collapsed")


def html(s: str):
    # strip indentation/blank lines so Markdown never turns HTML into a code block
    st.markdown("\n".join(l.strip() for l in s.splitlines() if l.strip()), unsafe_allow_html=True)


def secret(name: str, default: str = "") -> str:
    """Settings live on the server (Streamlit 'Secrets') and are never shown to visitors."""
    try:
        return str(st.secrets.get(name, os.environ.get(name, default)))
    except Exception:
        return os.environ.get(name, default)


# ------------------------------------------------------------------ styles (Vigyapan Studio palette)
html("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,600;12..96,800&family=DM+Sans:wght@400;500;700&family=DM+Mono:wght@400;500&display=swap');
:root{--bg:#06201F;--bg2:#0B2E2C;--panel:#0E3634;--line:#1E5551;--ink:#F3F7F2;--mute:#9DBFB9;
  --gold:#FFB72B;--gold-ink:#2A1A00;--coral:#FF5A45;--mint:#5EE6B8;
  --display:'Bricolage Grotesque','Arial Black',system-ui,sans-serif;--body:'DM Sans',system-ui,sans-serif;
  --mono:'DM Mono',ui-monospace,Menlo,monospace}
html,body,.stApp,p,label,input,button,textarea,select,li,span,div{font-family:var(--body)}
.stApp{background:var(--bg);color:var(--ink)}
header[data-testid="stHeader"],[data-testid="stToolbar"],[data-testid="stDecoration"],#MainMenu,footer,
[data-testid="stSidebar"],[data-testid="collapsedControl"]{display:none !important}
.block-container{max-width:100% !important;padding:0 !important}
.main [data-testid="stVerticalBlock"],[data-testid="stMainBlockContainer"] > div > [data-testid="stVerticalBlock"]{gap:0}
.stMarkdown p{margin:0}
a,a:hover,a:visited,.stMarkdown a{text-decoration:none !important;color:inherit}
.logo,.logo:hover{color:var(--ink) !important}
.links a.l,.foot a.f{color:var(--mute) !important}.links a.l:hover,.foot a.f:hover{color:var(--gold) !important}
.wrap{max-width:1120px;margin:0 auto;padding:0 24px;position:relative;z-index:1}
.aura{position:fixed;inset:0;pointer-events:none;z-index:0;overflow:hidden}
.aura i{position:absolute;border-radius:50%;filter:blur(80px);opacity:.33;animation:drift 18s ease-in-out infinite alternate}
.aura i:nth-child(1){width:440px;height:440px;background:#0F8C7E;top:-140px;left:-120px}
.aura i:nth-child(2){width:380px;height:380px;background:#FF5A45;opacity:.16;top:40%;right:-160px;animation-duration:22s}
.aura i:nth-child(3){width:340px;height:340px;background:#FFB72B;opacity:.12;bottom:-140px;left:25%;animation-duration:26s}
@keyframes drift{to{transform:translate(60px,40px) scale(1.15)}}
.mono{font-family:var(--mono)}

/* nav */
.nav{position:sticky;top:0;z-index:50;background:rgba(6,32,31,.88);backdrop-filter:blur(12px);border-bottom:1px solid var(--line)}
.nav .wrap{display:flex;align-items:center;justify-content:space-between;height:72px}
.logo{display:flex;align-items:center;gap:11px;font-family:var(--display);font-weight:800;font-size:1.2rem;color:var(--ink);letter-spacing:-.01em;white-space:nowrap}
.logo small{display:block;font-family:var(--mono);font-size:.62rem;font-weight:500;color:var(--mint);letter-spacing:.14em;margin-top:1px}
.links{display:flex;align-items:center;gap:28px}
.links a.l{color:var(--mute);font-weight:500;font-size:.94rem}.links a.l:hover{color:var(--gold)}
@media(max-width:820px){.links a.l{display:none}}
@media(max-width:480px){.nav .logo small{display:none}.nav .btn.sm{padding:8px 12px;font-size:.82rem}.logo{font-size:1.05rem}}

/* buttons (same 3D gold button as the Vigyapan Studio site) */
.btn{font-family:var(--display);font-weight:800;font-size:1.02rem;background:var(--gold);color:var(--gold-ink) !important;
  border-radius:14px;padding:14px 24px;display:inline-flex;align-items:center;gap:10px;position:relative;overflow:hidden;
  box-shadow:0 6px 0 #B57A00,0 14px 30px rgba(255,183,43,.22);transition:transform .12s,box-shadow .12s}
.btn:hover{transform:translateY(2px);box-shadow:0 4px 0 #B57A00,0 10px 24px rgba(255,183,43,.22)}
.btn::after{content:"";position:absolute;top:0;left:-60%;width:40%;height:100%;
  background:linear-gradient(100deg,transparent,rgba(255,255,255,.55),transparent);transform:skewX(-20deg);animation:shine 3.2s infinite}
@keyframes shine{0%,60%{left:-60%}100%{left:130%}}
.btn.sm{padding:10px 18px;font-size:.92rem;box-shadow:0 4px 0 #B57A00}
.btn.alt{background:transparent;color:var(--ink) !important;box-shadow:none;border:1.5px solid var(--line)}
.btn.alt::after{display:none}.btn.alt:hover{border-color:var(--gold);transform:none}

/* hero */
.hero{padding:72px 0 64px}
.hero .wrap{display:grid;grid-template-columns:1.1fr .9fr;gap:48px;align-items:center}
.pill{font-family:var(--mono);font-size:.72rem;letter-spacing:.08em;text-transform:uppercase;color:var(--mint);
  border:1px solid var(--line);padding:6px 12px;border-radius:99px;display:inline-flex;gap:8px;align-items:center}
.pill::before{content:"";width:7px;height:7px;border-radius:50%;background:var(--mint);animation:blink 1.6s infinite}
@keyframes blink{50%{opacity:.25}}
.eyebrow{font-family:var(--mono);font-size:.78rem;letter-spacing:.1em;text-transform:uppercase;color:var(--gold)}
.hero h1{font-family:var(--display);font-weight:800;font-size:clamp(2.4rem,5.6vw,4.2rem);line-height:.98;letter-spacing:-.03em;
  margin:16px 0 18px;color:var(--ink);padding:0}
.hero h1 em{font-style:normal;color:var(--gold);position:relative;white-space:nowrap}
.hero h1 em svg{position:absolute;left:0;bottom:-.14em;width:100%;height:.32em;overflow:visible}
.hero h1 em path{stroke:var(--coral);stroke-width:5;fill:none;stroke-linecap:round;stroke-dasharray:300;stroke-dashoffset:300;animation:draw 1.2s .5s ease-out forwards}
@keyframes draw{to{stroke-dashoffset:0}}
.lede{color:var(--mute);font-size:1.12rem;line-height:1.6;max-width:34em}
.cta{display:flex;gap:16px;margin:28px 0 24px;flex-wrap:wrap;align-items:center}
.trust{display:flex;flex-wrap:wrap;gap:8px 22px;color:var(--mute);font-size:.92rem}
.trust span::before{content:"✓ ";color:var(--mint);font-weight:700}
@media(max-width:900px){.hero .wrap{grid-template-columns:1fr}.mock{transform:none !important}}

/* mock report card */
.mock{background:var(--bg2);border:1px solid var(--line);border-radius:26px;padding:22px;
  box-shadow:0 30px 70px rgba(0,0,0,.5),0 0 0 8px rgba(255,255,255,.03);transform:rotate(2deg);animation:float 6s ease-in-out infinite}
@keyframes float{50%{transform:rotate(1deg) translateY(-10px)}}
.mock .top{display:flex;justify-content:space-between;align-items:center;border-bottom:1px dashed var(--line);padding-bottom:14px}
.mock .co{font-family:var(--display);font-weight:800;font-size:1.15rem}
.mock .co small{display:block;color:var(--mute);font-family:var(--mono);font-weight:400;font-size:.7rem;margin-top:2px}
.badge{font-family:var(--mono);font-size:.7rem;font-weight:500;padding:5px 10px;border-radius:99px;letter-spacing:.06em;text-transform:uppercase}
.b-hot{background:rgba(255,90,69,.18);color:#FFB3A8;border:1px solid rgba(255,90,69,.4)}
.b-warm{background:rgba(255,183,43,.16);color:#FFD98A;border:1px solid rgba(255,183,43,.4)}
.b-low{background:rgba(94,230,184,.14);color:#A5F3D6;border:1px solid rgba(94,230,184,.35)}
.b-check{background:rgba(157,191,185,.12);color:var(--mute);border:1px solid var(--line)}
.meter{margin-top:15px}.meter .h{display:flex;justify-content:space-between;font-family:var(--mono);font-size:.74rem;color:var(--mute);text-transform:uppercase;letter-spacing:.06em}
.meter .h b{font-weight:500}
.meter .bar{height:8px;border-radius:99px;background:var(--panel);margin-top:7px;overflow:hidden}
.meter .bar i{display:block;height:100%;border-radius:99px}
.chips{margin-top:16px;display:flex;flex-wrap:wrap;gap:6px}
.chip{font-family:var(--mono);font-size:.7rem;background:var(--panel);border:1px solid var(--line);color:var(--ink);padding:5px 9px;border-radius:8px}
.pitchq{margin-top:14px;background:rgba(255,183,43,.08);border:1.5px dashed var(--gold);border-radius:14px;padding:12px 14px;font-size:.88rem;line-height:1.5}
.mock .tag{font-family:var(--mono);font-size:.64rem;color:var(--mute);text-align:right;margin-top:10px;letter-spacing:.06em;text-transform:uppercase}

/* ticker */
.ticker{border-block:1px solid var(--line);overflow:hidden;white-space:nowrap;padding:12px 0;font-family:var(--mono);
  font-size:.8rem;color:var(--mute);text-transform:uppercase;letter-spacing:.1em;position:relative;z-index:1}
.ticker div{display:inline-block;animation:tick 30s linear infinite}
.ticker span{margin-inline:22px}.ticker span::before{content:"✦";color:var(--gold);margin-right:22px}
@keyframes tick{to{transform:translateX(-50%)}}

/* sections */
.sec{padding:76px 0 10px;position:relative;z-index:1}
.sec h2{font-family:var(--display);font-weight:800;font-size:clamp(1.8rem,4vw,2.7rem);line-height:1.05;letter-spacing:-.025em;
  margin:10px 0 28px;color:var(--ink);padding:0;max-width:22em}
.grid3{display:grid;grid-template-columns:repeat(3,1fr);gap:16px}
.grid4{display:grid;grid-template-columns:repeat(auto-fit,minmax(230px,1fr));gap:16px}
@media(max-width:860px){.grid3{grid-template-columns:1fr}}
.card{background:var(--bg2);border:1px solid var(--line);border-radius:20px;padding:24px;transition:transform .2s,border-color .2s}
.card:hover{transform:translateY(-6px) rotate(-.6deg);border-color:var(--gold)}
.card h3{font-family:var(--display);font-size:1.22rem;margin:14px 0 6px;color:var(--ink);padding:0;letter-spacing:-.01em}
.card p{color:var(--mute);font-size:.95rem;line-height:1.55}
.card svg{width:44px;height:44px}
.num{width:38px;height:38px;border-radius:50%;display:grid;place-items:center;font-family:var(--display);font-weight:800;
  background:var(--gold);color:var(--gold-ink)}

/* scanner */
.st-key-scanner{max-width:1072px;width:calc(100% - 48px);margin:0 auto 20px;background:var(--bg2);border:1px solid var(--line);
  border-radius:26px;padding:34px 36px 30px;position:relative;z-index:1}
.st-key-scanner [data-testid="stVerticalBlock"]{gap:.95rem}
@media(max-width:600px){.st-key-scanner{padding:22px 18px}}
.panel-h{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap;border-bottom:1px dashed var(--line);padding-bottom:18px}
.panel-h h3{font-family:var(--display);font-size:1.5rem;font-weight:800;margin:0;padding:0;color:var(--ink)}
.steps{display:flex;align-items:center;gap:10px;font-family:var(--mono);font-size:.72rem;letter-spacing:.06em;text-transform:uppercase;color:var(--mute);flex-wrap:wrap}
.steps .s{display:flex;align-items:center;gap:8px}
.steps .n{width:26px;height:26px;border-radius:50%;border:1.5px solid var(--line);display:grid;place-items:center}
.steps .s.on{color:var(--ink)}.steps .s.on .n{background:var(--gold);border-color:var(--gold);color:var(--gold-ink);font-weight:700}
.steps .s.done .n{background:var(--mint);border-color:var(--mint);color:#032A1E}
.steps .bar{width:26px;height:2px;background:var(--line)}
.sh{display:flex;align-items:center;gap:12px;margin-top:8px}
.sh .n{width:28px;height:28px;border-radius:50%;background:var(--gold);color:var(--gold-ink);display:grid;place-items:center;font-weight:800;font-size:.85rem}
.sh b{font-family:var(--display);font-size:1.15rem;color:var(--ink)}
.sh .s{margin-left:auto;font-family:var(--mono);font-size:.74rem;color:var(--mute);text-transform:uppercase;letter-spacing:.06em}
.note{background:var(--bg);border:1px solid var(--line);border-radius:14px;padding:13px 16px;color:var(--mute);font-size:.93rem}
.note b{color:var(--ink)}
.ok{background:rgba(94,230,184,.08);border:1px solid rgba(94,230,184,.35);border-radius:14px;padding:13px 16px;color:var(--ink);font-size:.95rem}
.tag2{display:inline-block;font-family:var(--mono);font-size:.72rem;border:1px solid var(--line);color:var(--mint);border-radius:8px;padding:2px 8px;margin:2px}

/* streamlit widgets -> dark theme */
.st-key-scanner label,.st-key-scanner [data-testid="stWidgetLabel"] p{color:var(--ink) !important;font-weight:700}
[data-testid="stFileUploaderDropzone"]{background:var(--bg) !important;border:2px dashed var(--line) !important;border-radius:18px !important;padding:26px !important}
[data-testid="stFileUploaderDropzone"]:hover{border-color:var(--gold) !important}
[data-testid="stFileUploaderDropzone"] *{color:var(--mute) !important}
[data-testid="stFileUploaderDropzone"] button{background:var(--panel) !important;border:1px solid var(--line) !important;color:var(--ink) !important}
[data-testid="stFileUploaderFile"]{background:var(--panel) !important;border:1px solid var(--line);border-radius:10px}
[data-testid="stFileUploaderFile"] *{color:var(--ink) !important}
[data-baseweb="input"] button,[data-baseweb="input"] > div{background:var(--bg) !important;color:var(--mute) !important}
[data-baseweb="input"]:focus-within{border-color:var(--gold) !important}
[data-testid="stTextInputRootElement"]{background:var(--bg) !important;border:1.5px solid var(--line) !important;border-radius:12px !important}
[data-testid="stTextInputRootElement"]:focus-within{border-color:var(--gold) !important}
[data-testid="stTextInputRootElement"] *{background:var(--bg) !important;border:0 !important}
[data-testid="stTextInputRootElement"] button{color:var(--mute) !important}
[data-testid="stFileChip"]{background:var(--panel) !important;border:1px solid var(--line) !important}
[data-testid="stFileChip"] *{color:var(--ink) !important}
.stTextInput input,[data-baseweb="select"] > div{background:var(--bg) !important;color:var(--ink) !important;border-color:var(--line) !important;border-radius:12px !important}
[data-baseweb="input"],[data-baseweb="base-input"]{background:var(--bg) !important;border-color:var(--line) !important;border-radius:12px !important}
[data-baseweb="select"] *{color:var(--ink) !important}
[data-baseweb="popover"] ul,[data-baseweb="popover"] li{background:var(--panel) !important;color:var(--ink) !important}
[data-baseweb="popover"] li:hover{background:var(--line) !important}
[data-testid="stExpander"] details{background:var(--bg) !important;border:1px solid var(--line) !important;border-radius:14px !important}
[data-testid="stExpander"] summary,[data-testid="stExpander"] summary *{color:var(--mute) !important}
.stButton button,.stDownloadButton button{font-family:var(--display) !important;font-weight:800 !important;font-size:1.05rem !important;
  border-radius:14px !important;padding:.85rem 1.4rem !important;transition:transform .12s,box-shadow .12s}
.stButton button p,.stDownloadButton button p{font-family:var(--display) !important;font-weight:800 !important;font-size:1.05rem !important}
button[kind="primary"]{background:var(--gold) !important;color:var(--gold-ink) !important;border:0 !important;
  box-shadow:0 6px 0 #B57A00,0 14px 30px rgba(255,183,43,.22) !important}
button[kind="primary"]:hover{transform:translateY(2px);box-shadow:0 4px 0 #B57A00 !important}
button[kind="primary"] p{color:var(--gold-ink) !important}
button[kind="secondary"]{background:transparent !important;border:1.5px solid var(--line) !important;color:var(--ink) !important}
button[kind="secondary"]:hover{border-color:var(--gold) !important}
button[kind="secondary"] p{color:var(--ink) !important}
[data-testid="stProgress"] p{color:var(--mute) !important;font-family:var(--mono) !important;font-size:.8rem !important}
[data-testid="stProgress"] div[role="progressbar"] > div{background:var(--panel) !important}
[data-testid="stProgress"] div[role="progressbar"] > div > div{background:var(--gold) !important}
[data-testid="stAlert"]{border-radius:14px}

/* results */
.tiles{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin:6px 0 16px}
@media(max-width:640px){.tiles{grid-template-columns:1fr 1fr}}
.tile{background:var(--bg);border:1px solid var(--line);border-radius:16px;padding:16px 18px}
.tile small{display:block;font-family:var(--mono);font-size:.68rem;color:var(--mute);letter-spacing:.06em;text-transform:uppercase}
.tile strong{display:block;font-family:var(--display);font-size:2rem;font-weight:800;line-height:1.1;margin-top:6px;color:var(--c)}
.tbl{margin-bottom:16px;max-height:460px;overflow:auto;border:1px solid var(--line);border-radius:16px}
.tbl table{width:100%;border-collapse:collapse;font-size:.88rem;min-width:720px}
.tbl th{position:sticky;top:0;background:var(--panel);color:var(--mute);font-family:var(--mono);font-weight:500;font-size:.7rem;
  letter-spacing:.06em;text-transform:uppercase;text-align:left;padding:11px 12px;border-bottom:1px solid var(--line)}
.tbl td{padding:11px 12px;border-bottom:1px solid rgba(30,85,81,.6);color:var(--ink);vertical-align:top}
.tbl td.m{color:var(--mute)}
.tbl tr:hover td{background:rgba(255,255,255,.02)}

/* faq + footer */
details.q{border-bottom:1px solid var(--line);padding:16px 0}
details.q summary{cursor:pointer;font-family:var(--display);font-weight:600;font-size:1.12rem;list-style:none;display:flex;justify-content:space-between;gap:14px;color:var(--ink)}
details.q summary::-webkit-details-marker{display:none}
details.q summary::after{content:"+";color:var(--gold);font-size:1.5rem;line-height:1;transition:transform .2s}
details.q[open] summary::after{transform:rotate(45deg)}
details.q p{color:var(--mute);margin:10px 0 0 !important;max-width:62ch;line-height:1.6}
.foot{padding:60px 0 28px;position:relative;z-index:1;border-top:1px solid var(--line);margin-top:70px}
.foot .wrap{display:grid;grid-template-columns:1.5fr 1fr 1fr;gap:28px}
.foot h4{font-family:var(--mono);font-size:.72rem;color:var(--gold);letter-spacing:.1em;text-transform:uppercase;margin:0 0 12px;padding:0}
.foot a.f{display:block;color:var(--mute);font-size:.92rem;margin:7px 0}.foot a.f:hover{color:var(--gold)}
.foot p{color:var(--mute);font-size:.92rem;line-height:1.6;max-width:330px;margin-top:14px !important}
.foot .bar{grid-column:1/-1;border-top:1px solid var(--line);padding-top:20px;margin-top:16px;display:flex;justify-content:space-between;
  flex-wrap:wrap;gap:10px;font-family:var(--mono);font-size:.76rem;color:var(--mute)}
.foot .bar b{color:var(--gold);font-weight:500}
@media(max-width:760px){.foot .wrap{grid-template-columns:1fr}}
@media (prefers-reduced-motion:reduce){*,*::before,*::after{animation-duration:.001s !important;animation-iteration-count:1 !important}}
</style>
<div class="aura" aria-hidden="true"><i></i><i></i><i></i></div>
""")

LOGO = """<svg width="40" height="40" viewBox="0 0 40 40" aria-hidden="true"><rect width="40" height="40" rx="12" fill="#FFB72B"/>
<path d="M16 12.5v15l12-7.5z" fill="#2A1A00"/><circle cx="31.5" cy="8.5" r="4" fill="#FF5A45" stroke="#06201F" stroke-width="2"/></svg>"""

# ------------------------------------------------------------------ nav + hero
html(f"""
<div class="nav"><div class="wrap">
  <a class="logo" href="#top">{LOGO}<div>{CLIENT_NAME}<small>LEAD SCANNER</small></div></a>
  <div class="links"><a class="l" href="#how">How it works</a><a class="l" href="#checks">What we check</a>
    <a class="l" href="#faq">FAQ</a><a class="btn sm" href="#scanner">Start scanning →</a></div>
</div></div>

<div class="hero" id="top"><div class="wrap">
  <div>
    <span class="pill">Scanner online · results in minutes</span>
    <div class="eyebrow" style="margin-top:18px">Lead scanner for event media</div>
    <h1>Turn every event contact into a <em>warm lead<svg viewBox="0 0 200 12" preserveAspectRatio="none"><path d="M3 8 C 40 2, 90 12, 130 6 S 185 4, 197 7"/></svg></em>.</h1>
    <p class="lede">Upload the contact list from any exhibition. We check every company's website, Google ranking and
      Instagram, then hand you a clear report of who needs what, with a ready pitch for each one.</p>
    <div class="cta"><a class="btn" href="#scanner">Scan my contact list <span aria-hidden="true">→</span></a>
      <a class="btn alt" href="#how">How it works</a></div>
    <div class="trust"><span>No setup needed</span><span>PDF + Excel report</span><span>Private to {CLIENT_NAME}</span></div>
  </div>
  <div class="mock" aria-hidden="true">
    <div class="top"><div class="co">Sharma Textiles<small>Anil Sharma · Ludhiana</small></div><span class="badge b-hot">Hot lead</span></div>
    <div class="meter"><div class="h"><span>Website</span><b style="color:#FFB3A8">28 / 100</b></div><div class="bar"><i style="width:28%;background:#FF5A45"></i></div></div>
    <div class="meter"><div class="h"><span>Google / SEO</span><b style="color:#FFD98A">54 / 100</b></div><div class="bar"><i style="width:54%;background:#FFB72B"></i></div></div>
    <div class="meter"><div class="h"><span>Instagram</span><b style="color:#FFB3A8">Not found</b></div><div class="bar"><i style="width:5%;background:#FF5A45"></i></div></div>
    <div class="chips"><span class="chip">Website redesign</span><span class="chip">SEO</span><span class="chip">Instagram setup</span></div>
    <div class="pitchq">“Website looks dated (not mobile-friendly, last updated 2016) and there's no Instagram page.”</div>
    <div class="tag">Example · one contact from a report</div>
  </div>
</div></div>
<div class="ticker" aria-hidden="true"><div>""" + "".join(
    f"<span>{t}</span>" for t in ["Website check", "Mobile-friendly test", "Old design finder", "SEO score",
                                  "Instagram check", "HOT / WARM leads", "Ready pitch lines", "PDF report"] * 4) + """</div></div>
""")

# ------------------------------------------------------------------ how it works + checks
html(f"""
<div class="sec" id="how"><div class="wrap">
  <div class="eyebrow">How it works</div>
  <h2>From a contact list to a call list in three steps.</h2>
  <div class="grid3">
    <div class="card"><div class="num">1</div><h3>Upload your Excel</h3><p>Drop in the contact list from an exhibition or event. The right columns are found automatically.</p></div>
    <div class="card"><div class="num">2</div><h3>We scan every contact</h3><p>Each company's website, Google visibility and Instagram are checked one by one, in minutes.</p></div>
    <div class="card"><div class="num">3</div><h3>Download your report</h3><p>A clear PDF with a card for every contact: what's wrong, what to offer, and a ready pitch line.</p></div>
  </div>
</div></div>

<div class="sec" id="checks"><div class="wrap">
  <div class="eyebrow">What we check</div>
  <h2>A full digital health check for every contact.</h2>
  <div class="grid4">
    <div class="card"><svg viewBox="0 0 44 44" fill="none"><rect x="4" y="7" width="36" height="30" rx="8" stroke="#FFB72B" stroke-width="3"/><path d="M4 16 H40" stroke="#FFB72B" stroke-width="3"/><circle cx="10" cy="11.5" r="1.6" fill="#FF5A45"/><circle cx="15" cy="11.5" r="1.6" fill="#5EE6B8"/><path d="M13 29 L19 23 L24 27 L31 21" stroke="#5EE6B8" stroke-width="3" stroke-linecap="round" stroke-linejoin="round"/></svg>
      <h3>Website</h3><p>No website, broken or "coming soon" pages, old design, not mobile-friendly, not secure, slow.</p></div>
    <div class="card"><svg viewBox="0 0 44 44" fill="none"><circle cx="19" cy="19" r="12" stroke="#5EE6B8" stroke-width="3"/><path d="M28 28 L39 39" stroke="#5EE6B8" stroke-width="3.5" stroke-linecap="round"/><path d="M13 22 L17 17 L21 20 L25 14" stroke="#FFB72B" stroke-width="2.6" stroke-linecap="round" stroke-linejoin="round"/></svg>
      <h3>Google / SEO</h3><p>Page title, Google description, headings, sitemap, analytics - the basics Google needs to show them.</p></div>
    <div class="card"><svg viewBox="0 0 44 44" fill="none"><rect x="5" y="5" width="34" height="34" rx="10" stroke="#FF5A45" stroke-width="3"/><circle cx="22" cy="22" r="7.5" stroke="#FFB72B" stroke-width="3"/><circle cx="31.5" cy="12.5" r="2" fill="#5EE6B8"/></svg>
      <h3>Instagram</h3><p>Whether they have a page at all, and how active and well set up it is.</p></div>
    <div class="card"><svg viewBox="0 0 44 44" fill="none"><path d="M6 10 a5 5 0 0 1 5-5 h22 a5 5 0 0 1 5 5 v16 a5 5 0 0 1-5 5 H18 l-8 7 v-7 h0 a5 5 0 0 1-4-5z" stroke="#FFB72B" stroke-width="3" stroke-linejoin="round"/><path d="M14 15 H30 M14 21 H25" stroke="#5EE6B8" stroke-width="3" stroke-linecap="round"/></svg>
      <h3>Ready pitch</h3><p>HOT / WARM / LOW priority, the services each lead needs, and one line to open the call.</p></div>
  </div>
</div></div>

<div class="sec" id="scanner" style="padding-bottom:26px"><div class="wrap">
  <div class="eyebrow">The scanner</div>
  <h2 style="margin-bottom:0">Scan your contact list.</h2>
</div></div>
""")

# ------------------------------------------------------------------ the scanner (real tool)
PRIO_BADGE = {"HOT": "b-hot", "WARM": "b-warm", "LOW": "b-low", "CHECK": "b-check"}


def steps_bar(active: int) -> str:
    names = ["Upload", "Scan", "Download"]
    out = []
    for i, n in enumerate(names, 1):
        cls = "on" if i == active else ("done" if i < active else "")
        mark = "✓" if i < active else str(i)
        out.append(f'<div class="s {cls}"><span class="n">{mark}</span>{n}</div>')
    return '<div class="steps">' + '<div class="bar"></div>'.join(out) + "</div>"


with st.container(key="scanner"):
    PASSWORD = secret("APP_PASSWORD")
    locked = bool(PASSWORD) and not st.session_state.get("ok")
    have_results = st.session_state.get("results") is not None
    head = st.empty()

    df = None
    up = None
    if locked:
        head.markdown(f'<div class="panel-h"><h3>Lead Scanner</h3><span class="badge b-low">🔒 Private to {CLIENT_NAME}</span></div>',
                      unsafe_allow_html=True)
        html('<div class="note">This scanner is private. Please enter the password you were given.</div>')
        pw = st.text_input("Password", type="password")
        if st.button("Unlock the scanner", type="primary"):
            if hmac.compare_digest(pw.strip(), PASSWORD):
                st.session_state["ok"] = True
                st.rerun()
            else:
                st.error("Wrong password. Please check and try again.")
    else:
        html('<div class="sh"><span class="n">1</span><b>Upload your Excel file</b></div>')
        up = st.file_uploader("Drag your Excel file here, or click to browse", type=["xlsx", "xls", "csv"])
        if not up:
            html('<div class="note">Your file just needs columns like <b>Company, Name, Phone, Email, Website</b>. '
                 'An Instagram column is optional.</div>')
        else:
            if st.session_state.get("file_id") != up.file_id:
                st.session_state.pop("results", None)
                st.session_state["file_id"] = up.file_id
                have_results = False
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

    if df is not None:
        MAX_ROWS = int(secret("MAX_ROWS", "3000"))
        if len(df) > MAX_ROWS:
            st.warning(f"This file has {len(df)} contacts. Only the first {MAX_ROWS} will be scanned - "
                       "please split bigger lists into parts.")
            df = df.head(MAX_ROWS)
        guess = auditor.detect_columns(df.columns)
        found = "".join(f'<span class="tag2">{auditor.FIELD_LABELS[f]}</span>' for f, c in guess.items() if c)
        html(f'<div class="ok">✓ Found <b>{len(df)} contacts</b>. Columns we understood: {found or "none"}</div>')

        with st.expander("Columns look wrong? Fix them here", expanded=not guess["website"] and not guess["email"]):
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
        else:
            workers = int(secret("WORKERS", "10"))
            mins = max(1, round(len(df) * 4 / workers / 60))
            html(f'<div class="sh"><span class="n">2</span><b>Start the scan</b><span class="s">≈ {mins} min</span></div>')
            if st.button("Start scan →", type="primary", width="stretch"):
                bar = st.progress(0.0, text="Starting…")
                note = st.empty()

                def progress(done, total, label):
                    bar.progress(done / total, text=f"Checked {done} of {total} · {label}")

                opts = {"apify_token": secret("APIFY_TOKEN"), "pagespeed_key": secret("PAGESPEED_KEY"),
                        "domain_age": secret("DOMAIN_AGE", "yes").lower() != "no", "workers": workers}
                res = auditor.scan_all(df, mapping, opts, progress, lambda m: note.info(m))
                bar.progress(1.0, text="Preparing your PDF report…")
                st.session_state["results"] = res
                st.session_state["pdf"] = build_pdf(res, up.name)
                st.session_state["report"] = build_report(res, up.name)
                st.session_state["src"] = up.name
                have_results = True
                bar.progress(1.0, text="Done - your report is ready below.")
                note.empty()

    res = st.session_state.get("results")
    if not locked:
        stage = 3 if (have_results and df is not None) else (2 if df is not None else 1)
        head.markdown(f'<div class="panel-h"><h3>Lead Scanner</h3>{steps_bar(stage)}</div>', unsafe_allow_html=True)

    if res is not None and df is not None and have_results:
        html('<div class="sh"><span class="n">3</span><b>Download your report</b></div>')
        base = os.path.splitext(st.session_state.get("src", "list"))[0]
        stamp = f"{dt.datetime.now():%d-%b-%Y}"
        c1, c2 = st.columns([3, 2])
        c1.download_button("Download PDF report", st.session_state["pdf"], type="primary", width="stretch",
                           file_name=f"{base} - Lead Report {stamp}.pdf", mime="application/pdf")
        c2.download_button("Download Excel", st.session_state["report"], width="stretch",
                           file_name=f"{base} - Lead Report {stamp}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        opp = res["Opportunities"].fillna("")
        tiles = [("Contacts checked", len(res), "#F3F7F2"),
                 ("HOT leads", int((res["Lead Priority"] == "HOT").sum()), "#FF5A45"),
                 ("No website", int(opp.str.contains("New Website").sum()), "#FFB72B"),
                 ("Website redesign / upgrade", int(opp.str.contains("Redesign|Upgrade").sum()), "#FFB72B"),
                 ("Need SEO", int(opp.str.contains("SEO").sum()), "#5EE6B8"),
                 ("Instagram leads", int(opp.str.contains("Instagram").sum()), "#5EE6B8")]
        html('<div class="tiles">' + "".join(f'<div class="tile" style="--c:{c}"><small>{l}</small><strong>{v}</strong></div>'
                                             for l, v, c in tiles) + "</div>")
        order = res["Lead Priority"].map({"HOT": 0, "WARM": 1, "CHECK": 2, "LOW": 3})
        view = res.assign(_o=order).sort_values("_o", kind="stable")

        def cell(v):
            return "" if v is None or (isinstance(v, float) and pd.isna(v)) else escape(str(v))

        rows = "".join(
            f'<tr><td><span class="badge {PRIO_BADGE.get(r["Lead Priority"], "b-check")}">{cell(r["Lead Priority"])}</span></td>'
            f'<td><b>{cell(r.get("Company"))}</b><br><span style="color:#9DBFB9">{cell(r.get("Contact Person"))}</span></td>'
            f'<td class="m">{cell(r.get("Phone"))}</td><td>{cell(r.get("Website Status"))}</td>'
            f'<td class="m">{cell(r.get("Opportunities"))}</td></tr>'
            for _, r in view.iterrows())
        html('<div class="tbl"><table><thead><tr><th>Priority</th><th>Company</th><th>Phone</th><th>Website</th>'
             f'<th>Needs</th></tr></thead><tbody>{rows}</tbody></table></div>')
        html('<div class="note">Every contact has its own detailed card - with all problems found and a ready pitch line - in the PDF report.</div>')

# ------------------------------------------------------------------ FAQ + footer
html(f"""
<div class="sec" id="faq"><div class="wrap"><div style="max-width:860px">
  <div class="eyebrow">Questions</div>
  <h2>Quick answers.</h2>
  <details class="q"><summary>What should my Excel file look like?</summary><p>Any normal contact list works. Columns like Company,
    Name, Phone, Email and Website are enough. The names don't need to be exact - "Mobile No." or "Email ID" are recognised too.
    An Instagram column is optional.</p></details>
  <details class="q"><summary>How long does a scan take?</summary><p>Roughly one minute for every 100 contacts. Keep the page
    open until your report is ready.</p></details>
  <details class="q"><summary>What do HOT, WARM and LOW mean?</summary><p>HOT needs two or more big services (for example no
    website and no Instagram) - call these first. WARM needs one or two. LOW is already in good shape. CHECK means the website
    blocked our automatic visit, so open it yourself.</p></details>
  <details class="q"><summary>Is my contact list safe?</summary><p>Your file is only used to run the scan and build your report.
    It isn't stored after you close the page.</p></details>
  <details class="q"><summary>How accurate is it?</summary><p>It's an automated first check of each company's home page. It's
    very good at spotting missing, broken and outdated websites - but always open a website yourself before you pitch.</p></details>
</div></div></div>

<div class="foot"><div class="wrap">
  <div><a class="logo" href="#top">{LOGO}<div>{CLIENT_NAME}<small>LEAD SCANNER</small></div></a>
    <p>Media coverage and podcasts for exhibitions and events across India - now with smarter lead finding.</p></div>
  <div><h4>Scanner</h4><a class="f" href="#how">How it works</a><a class="f" href="#checks">What we check</a><a class="f" href="#scanner">Start scanning</a></div>
  <div><h4>Help</h4><a class="f" href="#faq">FAQ</a><a class="f" href="#faq">Is my data safe?</a></div>
  <div class="bar"><span>© {dt.date.today().year} {CLIENT_NAME}. All rights reserved.</span><span>Built by <b>{POWERED_BY}</b></span></div>
</div></div>
""")
