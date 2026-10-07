"""
Vigyapan Lead Scanner - scanning engine.

Reads a contact database (Excel/CSV), checks every company's website,
SEO basics and social media presence, then scores each lead and tags
the services it needs (New Website, Redesign, SEO, Instagram ...).
"""
from __future__ import annotations

import datetime as dt
import json
import re
import socket
import ssl
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse

import pandas as pd
import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

THIS_YEAR = dt.date.today().year
TIMEOUT = 15
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/126.0 Safari/537.36")

# --------------------------------------------------------------------------
# 1. Reading the database & finding the right columns
# --------------------------------------------------------------------------
FIELDS = ["company", "name", "phone", "email", "website", "instagram", "city", "category"]
FIELD_LABELS = {
    "company": "Company", "name": "Contact Person", "phone": "Phone",
    "email": "Email", "website": "Website", "instagram": "Instagram",
    "city": "City", "category": "Category",
}
# Order matters: more specific fields are checked first.
COLUMN_HINTS = [
    ("email", ["email", "e-mail", "e mail", "mail id", "mail"]),
    ("instagram", ["instagram", "insta", "ig handle", "ig"]),
    ("website", ["website", "web site", "web", "url", "site", "domain", "webpage"]),
    ("phone", ["phone", "mobile", "contact no", "contact number", "whatsapp", "cell", "tel", "number", "mob"]),
    ("company", ["company", "organisation", "organization", "firm", "business", "exhibitor", "brand", "org"]),
    ("city", ["city", "location", "place", "state"]),
    ("category", ["category", "industry", "sector", "segment", "type of business"]),
    ("name", ["contact person", "person", "full name", "contact name", "name", "owner"]),
]


def _norm(s) -> str:
    return re.sub(r"[^a-z0-9 ]+", " ", str(s).lower()).strip()


def detect_columns(columns) -> dict:
    """Guess which spreadsheet column holds which field. Returns {field: column or None}."""
    mapping = {f: None for f in FIELDS}
    used = set()
    # exact match pass, then "contains" pass
    for mode in ("exact", "contains"):
        for field, hints in COLUMN_HINTS:
            if mapping[field]:
                continue
            for col in columns:
                if col in used:
                    continue
                n = _norm(col)
                ok = any(n == h for h in hints) if mode == "exact" else any(
                    re.search(rf"\b{re.escape(h)}\b", n) for h in hints)
                if ok:
                    mapping[field] = col
                    used.add(col)
                    break
    return mapping


def read_database(file, sheet_name=0) -> pd.DataFrame:
    name = getattr(file, "name", str(file)).lower()
    if name.endswith(".csv"):
        df = pd.read_csv(file, dtype=str)
    else:
        df = pd.read_excel(file, sheet_name=sheet_name, dtype=str)
    df = df.dropna(how="all").fillna("")
    df.columns = [str(c).strip() for c in df.columns]
    return df


# --------------------------------------------------------------------------
# 2. Small helpers
# --------------------------------------------------------------------------
EMPTY_VALUES = {"", "na", "n a", "n/a", "nil", "none", "-", "--", "0", "nan", "null", "no", "not available"}
FREE_EMAIL_DOMAINS = {
    "gmail.com", "googlemail.com", "yahoo.com", "yahoo.co.in", "yahoo.in", "ymail.com",
    "hotmail.com", "outlook.com", "live.com", "msn.com", "rediffmail.com", "rediff.com",
    "icloud.com", "me.com", "aol.com", "protonmail.com", "proton.me", "zoho.com",
    "zohomail.com", "mail.com", "gmx.com", "yandex.com", "inbox.com", "hotmail.co.in",
    "outlook.in", "sify.com", "vsnl.net", "vsnl.com",
}
TWO_LEVEL_TLDS = {"co.in", "org.in", "net.in", "firm.in", "gen.in", "ind.in", "ac.in", "edu.in",
                  "gov.in", "co.uk", "org.uk", "com.au", "net.au", "co.nz", "com.sg", "com.my",
                  "co.za", "com.br", "co.jp", "com.cn", "com.hk", "ae.org", "co.ae", "com.sa"}
SOCIAL_PATTERNS = {
    "Instagram": r"instagram\.com/",
    "Facebook": r"(facebook|fb)\.com/",
    "LinkedIn": r"linkedin\.com/",
    "YouTube": r"(youtube\.com/|youtu\.be/)",
    "X/Twitter": r"(twitter\.com/|//x\.com/)",
    "WhatsApp": r"(wa\.me/|api\.whatsapp\.com|whatsapp\.com/send)",
}
IG_RESERVED = {"p", "reel", "reels", "explore", "stories", "accounts", "tv", "about", "developer",
               "legal", "direct", "share", "sharer", "web"}
PARKED_PHRASES = ["domain is for sale", "buy this domain", "this domain may be for sale",
                  "domain for sale", "parked free", "parkingcrew", "sedoparking", "is parked",
                  "coming soon", "under construction", "site is under maintenance",
                  "website coming soon", "launching soon", "account suspended",
                  "this account has been suspended", "default web page", "index of /",
                  "future home of something quite cool", "welcome to nginx", "it works!"]
BUILDERS = [
    ("Wix", r"wix\.com|wixstatic\.com|X-Wix"),
    ("Shopify", r"cdn\.shopify\.com|myshopify\.com"),
    ("Squarespace", r"squarespace\.com|static1\.squarespace"),
    ("Blogger", r"blogger\.com|blogspot\.com"),
    ("Google Sites", r"sites\.google\.com"),
    ("Weebly", r"weebly\.com|editmysite\.com"),
    ("GoDaddy Builder", r"img1\.wsimg\.com|godaddy website builder"),
    ("WordPress", r"wp-content|wp-includes"),
    ("Joomla", r"/media/jui/|joomla"),
    ("Webflow", r"webflow\.com|assets\.website-files\.com"),
]


def is_empty(v) -> bool:
    return _norm(v) in EMPTY_VALUES


def normalize_url(raw) -> str | None:
    if raw is None or is_empty(raw):
        return None
    s = str(raw).strip().split()[0].strip(",;")
    if "@" in s and "/" not in s:          # someone typed an email in the website column
        return None
    if not re.match(r"^https?://", s, re.I):
        s = "http://" + s.lstrip("/")
    p = urlparse(s)
    if not p.netloc or "." not in p.netloc:
        return None
    return f"{p.scheme.lower()}://{p.netloc.lower()}{p.path or '/'}"


def registrable_domain(host: str) -> str:
    host = host.lower().split(":")[0]
    if host.startswith("www."):
        host = host[4:]
    parts = host.split(".")
    if len(parts) >= 3 and ".".join(parts[-2:]) in TWO_LEVEL_TLDS:
        return ".".join(parts[-3:])
    return ".".join(parts[-2:])


def is_ip(host: str) -> bool:
    return bool(re.match(r"^\d+\.\d+\.\d+\.\d+$", host.split(":")[0])) or host.startswith("localhost")


def email_domain(email) -> str | None:
    if is_empty(email):
        return None
    m = re.search(r"@([A-Za-z0-9.-]+\.[A-Za-z]{2,})", str(email))
    if not m:
        return None
    d = m.group(1).lower()
    return None if d in FREE_EMAIL_DOMAINS else d


def instagram_handle(value) -> str | None:
    """Pull a clean handle out of '@abc', 'abc', 'instagram.com/abc/?hl=en' ..."""
    if value is None or is_empty(value):
        return None
    s = str(value).strip()
    m = re.search(r"instagram\.com/([A-Za-z0-9_.]+)", s, re.I)
    if m:
        h = m.group(1)
    elif re.match(r"^@?[A-Za-z0-9_.]{2,30}$", s):
        h = s.lstrip("@")
    else:
        return None
    h = h.strip(".").lower()
    return None if h in IG_RESERVED or not h else h


# --------------------------------------------------------------------------
# 3. Network checks
# --------------------------------------------------------------------------
def _session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": UA, "Accept-Language": "en-IN,en;q=0.9",
                      "Accept": "text/html,application/xhtml+xml,*/*;q=0.8"})
    return s


def fetch_site(url: str, session: requests.Session) -> dict:
    """Try https first, then http. Returns info about what happened."""
    p = urlparse(url)
    candidates = [f"https://{p.netloc}{p.path}", f"http://{p.netloc}{p.path}"]
    last_error = ""
    for i, u in enumerate(candidates):
        for verify in (True, False):
            try:
                t0 = time.time()
                r = session.get(u, timeout=TIMEOUT, allow_redirects=True, verify=verify)
                elapsed = time.time() - t0
                return {"ok": True, "response": r, "elapsed": elapsed,
                        "ssl_valid": verify if u.startswith("https") else None,
                        "final_url": r.url}
            except requests.exceptions.SSLError as e:
                last_error = "SSL certificate problem"
                continue          # retry same URL without verification
            except requests.exceptions.Timeout:
                last_error = "Timed out (site too slow / not responding)"
                break
            except requests.exceptions.ConnectionError as e:
                msg = str(e)
                if "NameResolution" in msg or "getaddrinfo" in msg or "Name or service" in msg \
                        or "nodename nor servname" in msg:
                    last_error = "Domain does not exist / not pointing anywhere"
                else:
                    last_error = "Could not connect"
                break
            except Exception as e:  # noqa
                last_error = f"Error: {type(e).__name__}"
                break
    return {"ok": False, "error": last_error}


def ssl_days_left(host: str) -> int | None:
    try:
        ctx = ssl.create_default_context()
        with socket.create_connection((host, 443), timeout=8) as sock:
            with ctx.wrap_socket(sock, server_hostname=host) as ss:
                cert = ss.getpeercert()
        exp = dt.datetime.strptime(cert["notAfter"], "%b %d %H:%M:%S %Y %Z")
        return (exp - dt.datetime.utcnow()).days
    except Exception:
        return None


def domain_registered_year(domain: str, session: requests.Session) -> int | None:
    """Public RDAP lookup (free, no key) for the domain's registration date."""
    try:
        r = session.get(f"https://rdap.org/domain/{domain}", timeout=10)
        if r.status_code != 200:
            return None
        for ev in r.json().get("events", []):
            if ev.get("eventAction") == "registration" and ev.get("eventDate"):
                return int(ev["eventDate"][:4])
    except Exception:
        pass
    return None


def url_exists(url: str, session: requests.Session, must_contain: str | None = None) -> bool:
    try:
        r = session.get(url, timeout=10, allow_redirects=True, verify=False)
        if r.status_code != 200:
            return False
        if must_contain:
            return must_contain.lower() in r.text[:5000].lower()
        return True
    except Exception:
        return False


def pagespeed(url: str, api_key: str, session: requests.Session) -> dict:
    """Google PageSpeed Insights (optional, needs a free Google API key)."""
    try:
        r = session.get("https://www.googleapis.com/pagespeedonline/v5/runPagespeed",
                        params={"url": url, "strategy": "mobile", "key": api_key,
                                "category": ["performance", "seo"]}, timeout=90)
        cats = r.json()["lighthouseResult"]["categories"]
        return {"mobile_speed": round(cats["performance"]["score"] * 100),
                "google_seo": round(cats["seo"]["score"] * 100)}
    except Exception:
        return {}


# --------------------------------------------------------------------------
# 4. Analysing a website
# --------------------------------------------------------------------------
def analyse_html(html: str, final_url: str, headers) -> dict:
    soup = BeautifulSoup(html, "html.parser")
    low = html.lower()
    text = soup.get_text(" ", strip=True)
    info: dict = {}

    # --- basics
    title = (soup.title.string or "").strip() if soup.title and soup.title.string else ""
    md = soup.find("meta", attrs={"name": re.compile("^description$", re.I)})
    meta_desc = (md.get("content") or "").strip() if md else ""
    info["title"] = title
    info["meta_description"] = meta_desc
    info["h1_count"] = len(soup.find_all("h1"))
    imgs = soup.find_all("img")
    info["images"] = len(imgs)
    info["images_no_alt"] = sum(1 for i in imgs if not (i.get("alt") or "").strip())
    info["word_count"] = len(text.split())
    info["viewport"] = bool(soup.find("meta", attrs={"name": re.compile("^viewport$", re.I)}))
    info["canonical"] = any("canonical" in " ".join(l.get("rel") or []).lower() for l in soup.find_all("link"))
    info["open_graph"] = bool(soup.find("meta", attrs={"property": re.compile("^og:", re.I)}))
    info["schema"] = "application/ld+json" in low or "itemtype=\"http://schema.org" in low \
        or "itemtype=\"https://schema.org" in low
    info["analytics"] = any(k in low for k in ["googletagmanager.com", "google-analytics.com",
                                                "gtag(", "fbq(", "connect.facebook.net",
                                                "clarity.ms", "hotjar"])
    rm = soup.find("meta", attrs={"name": re.compile("^robots$", re.I)})
    info["noindex"] = bool(rm and "noindex" in (rm.get("content") or "").lower())
    info["lang"] = bool(soup.html and soup.html.get("lang"))
    info["favicon"] = any("icon" in " ".join(l.get("rel") or []).lower() for l in soup.find_all("link"))

    # --- age signals
    years = []
    for m in re.finditer(r"(?:©|&copy;|copyright|\(c\))\s*(?:[^0-9<]{0,40})?((?:19|20)\d{2})"
                         r"(?:\s*[-–—]\s*((?:19|20)\d{2}))?", html, re.I):
        years += [int(y) for y in m.groups() if y]
    years = [y for y in years if 1995 <= y <= THIS_YEAR + 1]
    info["copyright_year"] = max(years) if years else None
    lm = headers.get("Last-Modified")
    try:
        info["last_modified_year"] = dt.datetime.strptime(lm[5:16], "%d %b %Y").year if lm else None
    except Exception:
        info["last_modified_year"] = None

    # --- technology signals
    old = []
    if ".swf" in low or "shockwave-flash" in low:
        old.append("Flash")
    if re.search(r"<(font|marquee|center|blink)\b", low):
        old.append("1990s-style HTML tags")
    if "<frameset" in low:
        old.append("Frames")
    jq = re.search(r"jquery[.-]?(\d)\.(\d+)", low)
    if jq and int(jq.group(1)) < 3:
        old.append(f"Old jQuery {jq.group(1)}.{jq.group(2)}")
    bs = re.search(r"bootstrap[/@.-]?v?(\d)\.\d", low)
    if bs and int(bs.group(1)) < 4:
        old.append(f"Old Bootstrap {bs.group(1)}")
    gen = soup.find("meta", attrs={"name": re.compile("^generator$", re.I)})
    gen_txt = (gen.get("content") or "") if gen else ""
    wp = re.search(r"wordpress\s+(\d+)\.(\d+)", gen_txt, re.I)
    if wp and (int(wp.group(1)) < 6):
        old.append(f"Old WordPress {wp.group(1)}.{wp.group(2)}")
    if len(soup.find_all("table")) >= 4 and not soup.find(["header", "nav", "section", "footer"]):
        old.append("Table-based layout")
    info["old_tech"] = old

    builder = ""
    blob = low + " " + " ".join(f"{k}:{v}" for k, v in headers.items()).lower()
    for name, pat in BUILDERS:
        if re.search(pat, blob, re.I):
            builder = name
            break
    info["built_with"] = builder or (gen_txt.split(" ")[0] if gen_txt else "")

    # --- parked / empty
    snippet = text[:3000].lower()
    info["parked"] = (any(p in snippet for p in PARKED_PHRASES) and info["word_count"] < 400) \
        or info["word_count"] < 25

    # --- links: socials, contact
    hrefs = [a.get("href", "") for a in soup.find_all("a")]
    joined = " ".join(hrefs).lower()
    socials = {k: bool(re.search(p, joined)) for k, p in SOCIAL_PATTERNS.items()}
    info["socials"] = socials
    info["instagram_from_site"] = None
    for h in hrefs:
        hd = instagram_handle(h) if "instagram.com" in (h or "").lower() else None
        if hd:
            info["instagram_from_site"] = hd
            break
    info["has_contact"] = "mailto:" in joined or "tel:" in joined or bool(
        re.search(r"(\+91[\s-]?)?[6-9]\d{9}", text))
    return info


def website_age_label(info: dict, domain_year: int | None) -> str:
    last = max([y for y in (info.get("copyright_year"), info.get("last_modified_year")) if y] or [0])
    if not last:
        return "Unknown"
    gap = THIS_YEAR - last
    if gap <= 1:
        return f"Recent (updated {last})"
    if gap <= 3:
        return f"Ageing (last sign of update {last})"
    return f"Old (last sign of update {last})"


# --------------------------------------------------------------------------
# 5. Scanning one lead
# --------------------------------------------------------------------------
def scan_lead(rec: dict, options: dict, cache: dict) -> dict:
    """rec has keys from FIELDS. Returns a flat result dict."""
    opt = {"domain_age": True, "pagespeed_key": "", **(options or {})}
    out = {FIELD_LABELS[f]: rec.get(f, "") for f in FIELDS}
    web_issues, seo_issues, ig_issues, notes = [], [], [], []

    url = normalize_url(rec.get("website"))
    guessed = False
    if not url:
        d = email_domain(rec.get("email"))
        if d:
            url = f"http://{d}/"
            guessed = True
    ig = instagram_handle(rec.get("instagram"))

    res = {
        "Website Checked": "", "Website Status": "", "Website Score": None,
        "Website Age": "", "SEO Score": None, "SEO Status": "",
        "Instagram Handle": ig or "", "Instagram Status": "",
        "Other Socials Found": "", "HTTPS Secure": "", "Mobile Friendly": "",
        "Load Time (s)": None, "Copyright Year": None, "Domain Registered": None,
        "SSL Days Left": None, "Built With": "", "Old Technology": "",
        "Page Title": "", "Meta Description": "", "Mobile Speed (Google)": None,
    }

    site = None
    if url:
        key = registrable_domain(urlparse(url).netloc) if not is_ip(urlparse(url).netloc) else url
        if key in cache:
            site = cache[key]
        else:
            site = _check_site(url, opt)
            cache[key] = site
        res["Website Checked"] = site.get("final_url") or url
        if guessed:
            notes.append(f"No website in the list - tried the email domain ({urlparse(url).netloc})")

    # ---------- website verdict
    if not url:
        res["Website Status"] = "No Website"
        web_score = 0
        web_issues.append("No website found (not in list, and email is a free Gmail/Yahoo-type address)")
    elif not site["ok"]:
        if guessed:
            res["Website Status"] = "No Website"
            res["Website Checked"] = ""
            web_issues.append("No website found (email domain does not open a website either)")
        else:
            res["Website Status"] = "Broken / Not Opening"
            web_issues.append(f"Website does not open: {site.get('error')}")
        web_score = 0
    else:
        info = site["info"]
        web_score = 100
        status = site["status"]
        if status >= 400:
            web_score -= 60
            web_issues.append(f"Website shows an error page (HTTP {status})")
        if info["parked"]:
            web_score -= 50
            web_issues.append("Site is empty / 'coming soon' / parked domain")
        if not site["https"]:
            web_score -= 20
            web_issues.append("Not secure - no HTTPS (browsers show 'Not Secure')")
        elif site["ssl_valid"] is False:
            web_score -= 20
            web_issues.append("SSL certificate is invalid / expired")
        elif site["ssl_days"] is not None and site["ssl_days"] < 15:
            web_score -= 5
            web_issues.append(f"SSL certificate expires in {site['ssl_days']} days")
        if not info["viewport"]:
            web_score -= 25
            web_issues.append("Not mobile-friendly (no mobile layout)")
        cy = info["copyright_year"]
        if cy and THIS_YEAR - cy >= 3:
            web_score -= 15
            web_issues.append(f"Looks outdated - footer still says {cy}")
        if info["old_tech"]:
            web_score -= min(30, 10 * len(info["old_tech"]))
            web_issues.append("Old technology: " + ", ".join(info["old_tech"]))
        if site["elapsed"] > 6:
            web_score -= 15
            web_issues.append(f"Very slow to load ({site['elapsed']:.1f}s)")
        elif site["elapsed"] > 3:
            web_score -= 8
            web_issues.append(f"Slow to load ({site['elapsed']:.1f}s)")
        if info["built_with"] in ("Blogger", "Google Sites", "Weebly", "GoDaddy Builder"):
            web_score -= 5
            web_issues.append(f"Built on a basic DIY builder ({info['built_with']})")
        if not info["has_contact"]:
            web_score -= 5
            web_issues.append("No clickable phone/email/WhatsApp on homepage")
        ps = site.get("pagespeed", {})
        if ps.get("mobile_speed") is not None and ps["mobile_speed"] < 50:
            web_score -= 10
            web_issues.append(f"Poor Google mobile speed score ({ps['mobile_speed']}/100)")
        web_score = max(0, web_score)

        if status >= 400 or info["parked"]:
            res["Website Status"] = "Parked / Empty / Error"
        elif web_score < 60:
            res["Website Status"] = "Outdated - Needs Redesign"
        elif web_score < 80:
            res["Website Status"] = "Needs Upgrade"
        else:
            res["Website Status"] = "Good / Modern"

        res.update({
            "Website Age": website_age_label(info, site.get("domain_year")),
            "HTTPS Secure": "Yes" if site["https"] and site["ssl_valid"] is not False else "No",
            "Mobile Friendly": "Yes" if info["viewport"] else "No",
            "Load Time (s)": round(site["elapsed"], 1),
            "Copyright Year": info["copyright_year"],
            "Domain Registered": site.get("domain_year"),
            "SSL Days Left": site.get("ssl_days"),
            "Built With": info["built_with"],
            "Old Technology": ", ".join(info["old_tech"]),
            "Page Title": info["title"][:150],
            "Meta Description": info["meta_description"][:250],
            "Mobile Speed (Google)": ps.get("mobile_speed"),
        })

        # ---------- SEO
        seo = 100
        t, d = info["title"], info["meta_description"]
        if not t:
            seo -= 15; seo_issues.append("No page title")
        elif len(t) < 10 or len(t) > 70:
            seo -= 5; seo_issues.append(f"Page title length not ideal ({len(t)} characters)")
        if not d:
            seo -= 15; seo_issues.append("No meta description (Google snippet)")
        elif len(d) < 50 or len(d) > 170:
            seo -= 5; seo_issues.append(f"Meta description length not ideal ({len(d)} characters)")
        if info["h1_count"] == 0:
            seo -= 10; seo_issues.append("No main heading (H1)")
        elif info["h1_count"] > 1:
            seo -= 3; seo_issues.append(f"{info['h1_count']} main headings (H1) - should be 1")
        if info["images"] and info["images_no_alt"] / info["images"] > 0.3:
            seo -= 10; seo_issues.append(f"{info['images_no_alt']} of {info['images']} images have no alt text")
        if info["word_count"] < 250:
            seo -= 10; seo_issues.append(f"Very little text on homepage ({info['word_count']} words)")
        if not site["sitemap"]:
            seo -= 10; seo_issues.append("No sitemap.xml")
        if not site["robots"]:
            seo -= 5; seo_issues.append("No robots.txt")
        if not info["canonical"]:
            seo -= 3; seo_issues.append("No canonical tag")
        if not info["open_graph"]:
            seo -= 5; seo_issues.append("No social-share preview tags (Open Graph)")
        if not info["schema"]:
            seo -= 5; seo_issues.append("No business schema markup")
        if not info["analytics"]:
            seo -= 5; seo_issues.append("No Google Analytics / visitor tracking")
        if info["noindex"]:
            seo -= 25; seo_issues.append("Page is hidden from Google (noindex)")
        if not info["lang"]:
            seo -= 2
        if not info["favicon"]:
            seo -= 2; seo_issues.append("No favicon (browser tab icon)")
        if ps.get("google_seo") is not None and ps["google_seo"] < 80:
            seo -= 5; seo_issues.append(f"Google Lighthouse SEO score {ps['google_seo']}/100")
        seo = max(0, seo)
        res["SEO Score"] = seo
        res["SEO Status"] = "Needs SEO Work" if seo < 70 else ("Basic SEO OK" if seo < 85 else "Good")

        # socials from site
        if not ig and info["instagram_from_site"]:
            ig = info["instagram_from_site"]
            res["Instagram Handle"] = ig
        found = [k for k, v in info["socials"].items() if v and k != "Instagram"]
        res["Other Socials Found"] = ", ".join(found)

    res["Website Score"] = web_score
    if not res["SEO Status"]:
        res["SEO Status"] = "No Website" if res["Website Status"] == "No Website" else "Could not check"

    # ---------- Instagram (basic; detailed check added later in batch)
    if ig:
        res["Instagram Status"] = "Found (not checked in detail)"
    else:
        res["Instagram Status"] = "No Instagram Found"
        ig_issues.append("No Instagram page found (not in list, not linked on website)")

    res["_web_issues"] = web_issues
    res["_seo_issues"] = seo_issues
    res["_ig_issues"] = ig_issues
    res["_notes"] = notes
    out.update(res)
    return out


def _check_site(url: str, opt: dict) -> dict:
    s = _session()
    f = fetch_site(url, s)
    if not f["ok"]:
        return {"ok": False, "error": f["error"]}
    r = f["response"]
    final = r.url
    p = urlparse(final)
    base = f"{p.scheme}://{p.netloc}"
    ctype = r.headers.get("Content-Type", "")
    html = r.text if ("html" in ctype or not ctype or "<html" in r.text[:500].lower()) else ""
    info = analyse_html(html, final, r.headers)
    site = {
        "ok": True, "status": r.status_code, "final_url": final, "elapsed": f["elapsed"],
        "https": final.startswith("https"), "ssl_valid": f["ssl_valid"], "info": info,
        "ssl_days": None, "domain_year": None, "pagespeed": {},
    }
    robots = _safe_text(s, base + "/robots.txt")
    site["robots"] = bool(robots.strip()) and "<html" not in robots[:500].lower()
    sitemap_found = site["robots"] and "sitemap:" in robots.lower()
    if not sitemap_found:
        for path in ("/sitemap.xml", "/sitemap_index.xml", "/wp-sitemap.xml"):
            sm = _safe_text(s, base + path).lower()
            if "<urlset" in sm or "<sitemapindex" in sm:
                sitemap_found = True
                break
    site["sitemap"] = sitemap_found
    host = p.netloc.split(":")[0]
    if site["https"] and not is_ip(host):
        site["ssl_days"] = ssl_days_left(host)
    if opt.get("domain_age") and not is_ip(host):
        site["domain_year"] = domain_registered_year(registrable_domain(host), s)
    if opt.get("pagespeed_key"):
        site["pagespeed"] = pagespeed(final, opt["pagespeed_key"], s)
    return site


_text_cache: dict = {}


def _safe_text(s, url) -> str:
    if url in _text_cache:
        return _text_cache[url]
    try:
        r = s.get(url, timeout=10, verify=False)
        t = r.text[:20000] if r.status_code == 200 else ""
    except Exception:
        t = ""
    _text_cache[url] = t
    return t


# --------------------------------------------------------------------------
# 6. Optional: detailed Instagram check via Apify (paid service, needs a token)
# --------------------------------------------------------------------------
def instagram_details_apify(handles: list[str], token: str, log=lambda m: None) -> dict:
    """Uses Apify's 'Instagram Profile Scraper' to fetch public profile data.
    Returns {handle: profile_dict}. Silently returns {} on any failure."""
    results = {}
    url = ("https://api.apify.com/v2/acts/apify~instagram-profile-scraper/"
           "run-sync-get-dataset-items")
    for i in range(0, len(handles), 40):
        chunk = handles[i:i + 40]
        try:
            r = requests.post(url, params={"token": token}, json={"usernames": chunk}, timeout=300)
            if r.status_code >= 300:
                log(f"Instagram check failed (Apify said {r.status_code}). Check your token / credit.")
                break
            for item in r.json():
                u = (item.get("username") or "").lower()
                if u:
                    results[u] = item
        except Exception as e:  # noqa
            log(f"Instagram check error: {e}")
            break
    return results


def apply_instagram_details(row: dict, prof: dict | None) -> None:
    handle = row.get("Instagram Handle")
    if not handle:
        return
    issues = row["_ig_issues"]
    if not prof or prof.get("error"):
        row["Instagram Status"] = "Handle not found on Instagram"
        issues.append(f"@{handle} does not exist or was removed - needs a new page")
        return
    followers = prof.get("followersCount") or 0
    posts = prof.get("postsCount") or 0
    row["Instagram Followers"] = followers
    row["Instagram Posts"] = posts
    latest = prof.get("latestPosts") or []
    last_date = None
    eng = None
    if latest:
        dates = [p.get("timestamp") for p in latest if p.get("timestamp")]
        if dates:
            last_date = max(dates)[:10]
        likes = [(p.get("likesCount") or 0) + (p.get("commentsCount") or 0) for p in latest[:12]]
        if followers and likes:
            eng = round(100 * sum(likes) / len(likes) / followers, 2)
    row["Instagram Last Post"] = last_date
    row["Instagram Engagement %"] = eng
    if prof.get("private"):
        issues.append("Account is private - customers can't see posts")
    if posts == 0:
        issues.append("No posts at all")
    elif posts < 30:
        issues.append(f"Only {posts} posts - looks inactive")
    if last_date:
        days = (dt.date.today() - dt.date.fromisoformat(last_date)).days
        if days > 30:
            issues.append(f"No post in {days} days")
    if followers < 1000:
        issues.append(f"Small audience ({followers} followers)")
    if not (prof.get("biography") or "").strip():
        issues.append("Empty bio")
    if not prof.get("externalUrl"):
        issues.append("No link in bio (website/WhatsApp)")
    if not (prof.get("isBusinessAccount") or prof.get("businessCategoryName")):
        issues.append("Not set up as a Business account")
    if eng is not None and eng < 1:
        issues.append(f"Low engagement ({eng}% per post)")
    row["Instagram Status"] = "Needs Work" if len(issues) >= 2 else ("Minor Fixes" if issues else "Good")


# --------------------------------------------------------------------------
# 7. Final verdict: what can we sell this lead?
# --------------------------------------------------------------------------
def finalise(row: dict) -> dict:
    opp = []
    ws = row["Website Status"]
    if ws in ("No Website", "Broken / Not Opening", "Parked / Empty / Error"):
        opp.append("New Website")
    elif ws == "Outdated - Needs Redesign":
        opp.append("Website Redesign")
    elif ws == "Needs Upgrade":
        opp.append("Website Upgrade")
    if row["SEO Status"] == "Needs SEO Work" or "New Website" in opp:
        opp.append("SEO")
    if row["Instagram Status"] in ("No Instagram Found", "Handle not found on Instagram"):
        opp.append("Instagram Page Setup")
    elif row["Instagram Status"] in ("Needs Work",):
        opp.append("Instagram Management")
    if ws not in ("No Website", "Broken / Not Opening") and not row.get("Other Socials Found") \
            and "Instagram Page Setup" in opp:
        opp.append("Social Media Setup")

    big = {"New Website", "Website Redesign", "Instagram Page Setup"}
    n_big = len(big.intersection(opp))
    if n_big >= 2 or (len(opp) >= 3):
        prio = "HOT"
    elif opp:
        prio = "WARM"
    else:
        prio = "LOW"

    row["Opportunities"] = ", ".join(opp) if opp else "None - digitally healthy"
    row["Lead Priority"] = prio
    row["Website Issues"] = "\n".join("• " + i for i in row.pop("_web_issues"))
    row["SEO Issues"] = "\n".join("• " + i for i in row.pop("_seo_issues"))
    row["Instagram Issues"] = "\n".join("• " + i for i in row.pop("_ig_issues"))
    row["Notes"] = "\n".join(row.pop("_notes"))
    row["Pitch Note"] = pitch_note(row, opp)
    return row


def pitch_note(row: dict, opp: list[str]) -> str:
    who = row.get("Company") or row.get("Contact Person") or "This business"
    bits = []
    ws = row["Website Status"]
    if ws == "No Website":
        bits.append("no website - people who search for them online find nothing")
    elif ws == "Broken / Not Opening":
        bits.append("website is not opening - every visitor is lost right now")
    elif ws == "Parked / Empty / Error":
        bits.append("website is empty / 'coming soon'")
    elif ws in ("Outdated - Needs Redesign", "Needs Upgrade"):
        issues = row["Website Issues"].lower()
        why = []
        for key, short in (("mobile", "not mobile-friendly"), ("footer still says", None),
                           ("old technology", "built on old technology"), ("https", "shows 'Not Secure'"),
                           ("slow", "slow to load")):
            if key in issues:
                if short is None:
                    m = re.search(r"footer still says (\d{4})", issues)
                    short = f"last updated around {m.group(1)}" if m else "looks old"
                why.append(short)
        bits.append("website looks dated (" + ", ".join(why[:2]) + ")" if why else "website needs a refresh")
    if "SEO" in opp and ws not in ("No Website", "Broken / Not Opening", "Parked / Empty / Error"):
        bits.append(f"SEO score only {row['SEO Score']}/100, so hard to find on Google")
    if "Instagram Page Setup" in opp:
        bits.append("no Instagram page")
    elif "Instagram Management" in opp:
        bits.append("Instagram is inactive / weak")
    if not bits:
        return "Digitally healthy - pitch event coverage / podcast feature instead."
    return f"{who}: " + "; ".join(bits) + "."


# --------------------------------------------------------------------------
# 8. Scan everything
# --------------------------------------------------------------------------
def scan_all(df: pd.DataFrame, mapping: dict, options: dict | None = None,
             progress=lambda done, total, label: None, log=lambda m: None) -> pd.DataFrame:
    options = options or {}
    records = []
    for idx, r in df.iterrows():
        rec = {f: (str(r[c]).strip() if c and c in df.columns else "") for f, c in mapping.items()}
        rec["_row"] = idx
        records.append(rec)

    cache: dict = {}
    results = [None] * len(records)
    workers = int(options.get("workers", 8))
    done = 0
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(_safe_scan, rec, options, cache): i for i, rec in enumerate(records)}
        for fut in as_completed(futs):
            i = futs[fut]
            results[i] = fut.result()
            done += 1
            progress(done, len(records), results[i].get("Company") or results[i].get("Contact Person") or "")

    token = options.get("apify_token")
    if token:
        handles = sorted({r["Instagram Handle"] for r in results if r.get("Instagram Handle")})
        if handles:
            log(f"Checking {len(handles)} Instagram profiles in detail...")
            profiles = instagram_details_apify(handles, token, log)
            if profiles:
                for r in results:
                    if r.get("Instagram Handle"):
                        apply_instagram_details(r, profiles.get(r["Instagram Handle"].lower()))

    final = [finalise(r) for r in results]
    out = pd.DataFrame(final)
    # keep any extra columns from the original sheet (e.g. event name, stall no.)
    used = {c for c in mapping.values() if c}
    extras = [c for c in df.columns if c not in used]
    for c in extras:
        out[f"[Original] {c}"] = df[c].values
    return out


def _safe_scan(rec, options, cache):
    try:
        return scan_lead(rec, options, cache)
    except Exception as e:  # never let one bad row stop the whole scan
        out = {FIELD_LABELS[f]: rec.get(f, "") for f in FIELDS}
        out.update({"Website Status": "Could not check", "Website Score": None,
                    "SEO Score": None, "SEO Status": "Could not check",
                    "Instagram Handle": instagram_handle(rec.get("instagram")) or "",
                    "Instagram Status": "Found (not checked in detail)" if instagram_handle(rec.get("instagram")) else "No Instagram Found",
                    "_web_issues": [f"Scanner error: {type(e).__name__}"], "_seo_issues": [],
                    "_ig_issues": [], "_notes": []})
        return out
