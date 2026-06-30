"""
OpportunitiesCorners.com Scraper  →  Google Sheets Version
============================================================
Site   : https://opportunitiescorners.com/category/bachelor-master-phd-scholarships/
Output : Google Sheets → Tab: oc_scholarships

NOTE: Same columns as BrightScholarship:
      "Offered By" + "Eligible Nations"
"""

import os
import re
import json
import asyncio
import time
from datetime import datetime, date

import gspread
from google.oauth2.service_account import Credentials
from playwright.async_api import async_playwright
from openai import AsyncOpenAI
from dotenv import load_dotenv

load_dotenv()

# ══════════════════════════════════════════════════════════════
# CONFIG
# ══════════════════════════════════════════════════════════════
BASE_URL      = "https://opportunitiescorners.com/category/bachelor-master-phd-scholarships/"
PAGE_URL_TMPL = "https://opportunitiescorners.com/category/bachelor-master-phd-scholarships/page/{n}/"
START_PAGE    = 1
MAX_PAGES     = 11
NAV_TIMEOUT   = 60_000
PAGE_WAIT     = 2500
DETAIL_WAIT   = 3000
TODAY         = date.today()

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS CONFIG
# ══════════════════════════════════════════════════════════════
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "oc_scholarships"

# Same columns as BrightScholarship
COLUMN_HEADERS = [
    "Scholarship Name", "Scholarship Link", "Offered By",
    "Official Website", "Scholarship Type", "Degree Type",
    "Host Country", "Eligible Nations", "Deadline", "Field of Study",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ══════════════════════════════════════════════════════════════
# OPENAI
# ══════════════════════════════════════════════════════════════
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
ai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

EMPTY = {
    "scholarship_type": "",
    "deadline":         "",
    "host_country":     "",
    "degree_type":      "",
    "field_of_study":   "",
}

# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS HELPERS
# ══════════════════════════════════════════════════════════════
_gs_client      = None
_gs_worksheet   = None
_header_written = False


def _get_worksheet() -> gspread.Worksheet:
    global _gs_client, _gs_worksheet, _header_written
    if _gs_worksheet is None:
        creds       = Credentials.from_service_account_file(CREDENTIALS_FILE, scopes=SCOPES)
        _gs_client  = gspread.authorize(creds)
        spreadsheet = _gs_client.open_by_key(SHEET_ID)
        try:
            _gs_worksheet = spreadsheet.worksheet(SHEET_TAB)
        except gspread.WorksheetNotFound:
            _gs_worksheet = spreadsheet.add_worksheet(title=SHEET_TAB, rows=10000, cols=20)
            _gs_worksheet.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
            _header_written = True
            print(f"  📋  Created tab '{SHEET_TAB}' with headers")
    return _gs_worksheet


def load_existing_links() -> set:
    global _header_written
    try:
        ws   = _get_worksheet()
        rows = ws.get_all_values()
        if not rows:
            ws.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
            _header_written = True
            print(f"  📝  Header row written to '{SHEET_TAB}'")
            return set()
        _header_written = True
        headers = rows[0]
        try:
            col_idx = headers.index("Scholarship Link")
        except ValueError:
            return set()
        return set(
            rows[i][col_idx].strip()
            for i in range(1, len(rows))
            if len(rows[i]) > col_idx and rows[i][col_idx].strip()
        )
    except Exception as e:
        print(f"  ⚠  Could not load existing links: {e}")
        return set()


def save_one_to_sheets(record: dict, retries: int = 3):
    global _header_written
    for attempt in range(retries):
        try:
            ws = _get_worksheet()
            if not _header_written:
                ws.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
                _header_written = True
                time.sleep(1)
            row = [str(record.get(col, "") or "") for col in COLUMN_HEADERS]
            ws.append_row(row, value_input_option="USER_ENTERED")
            return
        except gspread.exceptions.APIError as e:
            if "RESOURCE_EXHAUSTED" in str(e) or "429" in str(e):
                wait = (attempt + 1) * 15
                print(f"  ⏳  Rate limit — waiting {wait}s ...")
                time.sleep(wait)
            else:
                print(f"  ⚠  Sheets API error: {e}")
                return
        except Exception as e:
            print(f"  ⚠  save_one error: {e}")
            time.sleep(5)


# ══════════════════════════════════════════════════════════════
# JAVASCRIPT
# ══════════════════════════════════════════════════════════════
COLLECT_CARDS_JS = r"""
() => {
    const results = [], seen = new Set();
    document.querySelectorAll('a.td-image-wrap[href][title]').forEach(a => {
        const href  = (a.getAttribute('href')  || '').trim();
        const title = (a.getAttribute('title') || '').trim();
        if (!href || !title || seen.has(href)) return;
        if (!href.includes('opportunitiescorners.com')) return;
        const low = href.toLowerCase();
        const skip = ['internship','conference','online-course','fellowship-list','exchange-program','/jobs','/result','how-to','what-is'];
        for (const w of skip) { if (low.includes(w)) return; }
        seen.add(href);
        results.push({ title, link: href });
    });
    return results;
}
"""

EXTRACT_DETAIL_JS = r"""
() => {
    const fields = {};
    document.querySelectorAll('table tr, .wp-block-table tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'));
        if (cells.length >= 2) {
            const label = (cells[0].innerText || '').trim();
            const value = (cells[1].innerText || '').trim();
            if (label && value) fields[label.toLowerCase()] = value;
        }
    });
    const content = document.querySelector('.td-post-content, .entry-content, article');
    if (content) {
        content.querySelectorAll('li, p').forEach(el => {
            const txt = (el.innerText || '').trim();
            const m = txt.match(/^([A-Za-z][A-Za-z\s\/]{1,40}?)\s*:\s*(.+)$/);
            if (m) { const l = m[1].trim().toLowerCase(), v = m[2].trim(); if (l && v && !fields[l]) fields[l] = v; }
        });
        content.querySelectorAll('p, li').forEach(el => {
            el.querySelectorAll('strong, b').forEach(s => {
                const label = (s.innerText||'').replace(/:$/,'').trim().toLowerCase();
                let value = '', node = s.nextSibling;
                while (node) {
                    if (node.nodeType===3){const t=node.textContent.replace(/^[\s:]+/,'').trim();if(t){value=t;break;}}
                    else if(node.nodeType===1){const t=(node.innerText||'').trim();if(t){value=t;break;}}
                    node=node.nextSibling;
                }
                if (label && value && !fields[label]) fields[label] = value;
            });
        });
    }
    let officialUrl = '';
    const centered = document.querySelectorAll('p[style*="text-align: center"] a[target="_blank"], p[style*="text-align:center"] a[target="_blank"], div[style*="text-align: center"] a[target="_blank"]');
    for (const a of centered) { const h=(a.getAttribute('href')||'').trim(); if(h.startsWith('http')&&!h.includes('opportunitiescorners.com')){officialUrl=h;break;} }
    if (!officialUrl) {
        for (const a of document.querySelectorAll('a.td_btn[target="_blank"],a[class*="td-btn"][target="_blank"],a[class*="td_btn"][target="_blank"]')) {
            const h=(a.getAttribute('href')||'').trim(); if(h.startsWith('http')&&!h.includes('opportunitiescorners.com')){officialUrl=h;break;}
        }
    }
    if (!officialUrl && content) {
        for (const a of content.querySelectorAll('a[target="_blank"]')) {
            const h=(a.getAttribute('href')||'').trim(), t=(a.innerText||'').trim().toLowerCase();
            if(h.startsWith('http')&&!h.includes('opportunitiescorners.com')&&(t.includes('visit')||t.includes('apply')||t.includes('official')||t.includes('here'))){officialUrl=h;break;}
        }
    }
    const clone = content ? content.cloneNode(true) : document.body.cloneNode(true);
    clone.querySelectorAll('script,style,nav,header,footer,.sharethis-inline-share-buttons,.td-related-title,.td-post-next-prev,.td-author-box,[class*="comment"],[id*="ads"],[class*="adsbygoogle"]').forEach(n=>n.remove());
    const text = (clone.innerText||'').trim();
    return { fields, officialUrl, text };
}
"""

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def safe_str(val) -> str:
    if val is None: return ""
    if isinstance(val, (list, tuple)): return ", ".join(str(v).strip() for v in val if v)
    return str(val).strip()

def clean(s) -> str:
    s = safe_str(s)
    for old, new in {"\u2018":"'","\u2019":"'","\u201c":'"',"\u201d":'"',"\u2013":"-","\u2014":"-","\u2026":"...","\u00a0":" "}.items():
        s = s.replace(old, new)
    return re.sub(r"\s+", " ", s).strip()

def clean_degree(s: str) -> str:
    parts = [p.strip() for p in re.split(r"[,;/]+", s) if p.strip()]
    out, seen = [], set()
    for p in parts:
        p = re.sub(r"^Master'?s?$","Masters",p,flags=re.IGNORECASE)
        p = re.sub(r"^Bachelor'?s?$","Bachelors",p,flags=re.IGNORECASE)
        if p.lower() not in seen: seen.add(p.lower()); out.append(p)
    return ", ".join(out)

def parse_deadline(s: str) -> date | None:
    s = clean(s)
    if not s or s.lower() in ("none","n/a","-","open","ongoing",""): return None
    s = re.sub(r'(\d+)(st|nd|rd|th)\b', r'\1', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+(for|and)\s+.+$', '', s, flags=re.IGNORECASE).strip()
    for fmt in ["%d %B %Y","%d %b %Y","%B %d, %Y","%b %d, %Y","%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%d-%B-%Y","%d-%b-%Y","%d %B, %Y","%d %b, %Y","%B %Y","%b %Y"]:
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    m = re.search(r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)[,\s]+(\d{4})", s, re.I)
    if m:
        try: return datetime.strptime(f"{m.group(1)} {m.group(2).title()} {m.group(3)}", "%d %B %Y").date()
        except ValueError: pass
    m = re.search(r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})", s, re.I)
    if m:
        try: return datetime.strptime(f"{m.group(2)} {m.group(1).title()} {m.group(3)}", "%d %B %Y").date()
        except ValueError: pass
    return None

def deadline_passed(s: str) -> bool:
    d = parse_deadline(s)
    return d is not None and d < TODAY

def get_field(fields: dict, *keys) -> str:
    for key in keys:
        for fk, fv in fields.items():
            if key.lower() in fk.lower():
                val = clean(fv)
                if val: return val
    return ""


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(name, fields, text):
    table_rows = "\n".join(f"  {k}: {v}" for k, v in fields.items()) if fields else "  (none found)"
    return (
        "You are a precise scholarship data extraction engine.\n\n"
        "STRICT RULES:\n- Return ONLY valid JSON.\n- Plain ASCII strings only.\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated from: Bachelors, Masters, MBA, PhD, PostDoc.\n"
        "- scholarship_type: e.g. 'Fully Funded', 'Partial Funding', 'Tuition waiver'.\n"
        "- deadline: exact date string as found. '' if not found.\n"
        "- host_country: country where student will STUDY.\n"
        "- field_of_study: comma-separated subject fields if mentioned.\n\n"
        '{"scholarship_type":"","deadline":"","host_country":"","degree_type":"","field_of_study":""}\n\n'
        f"Name: {name}\n\nTable:\n{table_rows}\n\nText:\n{text[:7000]}"
    )

async def ai_extract(name, fields, text):
    try:
        resp = await ai_client.chat.completions.create(
            model="gpt-4o-mini", temperature=0,
            messages=[{"role":"system","content":"Scholarship data extraction. Return only valid JSON."},{"role":"user","content":build_prompt(name,fields,text)}],
            max_tokens=400,
        )
        raw = resp.choices[0].message.content or ""
    except Exception as e:
        print(f"  ⚠  OpenAI: {e}"); return dict(EMPTY)
    raw = re.sub(r"^```[a-z]*\s*","",raw.strip(),flags=re.I); raw = re.sub(r"\s*```$","",raw).strip()
    try:
        d = json.loads(raw); return d if isinstance(d, dict) else dict(EMPTY)
    except Exception:
        m = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if m:
            try: d = json.loads(m.group(0)); return d if isinstance(d, dict) else dict(EMPTY)
            except Exception: pass
    return dict(EMPTY)


# ══════════════════════════════════════════════════════════════
# SCRAPE ONE CARD
# ══════════════════════════════════════════════════════════════

async def scrape_one(page, card, seq, total):
    name = clean(card.get("title","")); link = card["link"]
    print(f"\n  ── ({seq}/{total}) {name[:58]}")
    try:
        await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await page.wait_for_timeout(DETAIL_WAIT)
    except Exception as e:
        print(f"     ⚠ Load failed: {e}"); return None

    data = await page.evaluate(EXTRACT_DETAIL_JS)
    fields = data.get("fields",{}); official_url = clean(data.get("officialUrl","")); text = clean(data.get("text",""))

    offered_by   = get_field(fields, "offered by","university name","host university","provider","organization","host")
    degree_raw   = get_field(fields, "degree level","degree","eligible degree","level")
    coverage     = get_field(fields, "scholarship coverage","financial benefits","coverage","funding","scholarship type","award")
    nationality  = get_field(fields, "eligible nationality","nationality","open to","eligible for")
    host_country = get_field(fields, "host country","award country","country","location","study in")
    deadline_raw = get_field(fields, "deadline","deadlines","last date","closing date","apply before","due date")

    if not deadline_raw and text:
        for pat in [
            r'(?:deadline|last date|closing date|apply before|due date|last day)\s*[:\-–]\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})',
            r'(?:deadline|last date|closing date|apply before|due date|last day)\s*[:\-–]\s*([A-Za-z]+\s+\d{1,2},?\s+\d{4})',
            r'Deadline[^\n]{0,30}?(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})',
            r'Last Date[^\n]{0,30}?(\d{1,2}\s+[A-Za-z]+\s+\d{4})',
        ]:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                c = m.group(1).strip().rstrip('.')
                if parse_deadline(c): deadline_raw = c; print(f"     deadline from text: {deadline_raw}"); break

    if not offered_by:
        for pat in [
            r'((?:University|Institute|College|School|Hochschule)\s+(?:of\s+)?[A-Z][a-zA-Z\s\-]{2,40}?)(?:\s+(?:Scholarship|Fellowship|2\d{3}|in\b))',
            r'([A-Z][a-zA-Z\s\-]{2,40}?\s+(?:University|Institute|College|School))(?:\s+(?:Scholarship|Fellowship|2\d{3}|in\b))',
            r'((?:Ministry|Government|Agency)\s+of\s+[A-Z][a-zA-Z\s]{2,40}?)(?:\s+(?:Scholarship|Fellowship|2\d{3}))',
        ]:
            m = re.search(pat, name)
            if m: offered_by = m.group(1).strip(); break

    if not offered_by and text:
        for pat in [r'(?:Offered by|Host University|Provider)[:\s]+([A-Z][^\n]{4,70}?)(?:\n|$)']:
            m = re.search(pat, text, re.IGNORECASE)
            if m:
                c = m.group(1).strip().rstrip('.')
                if 4 < len(c) < 80: offered_by = c; break

    print(f"     fields={len(fields)}  offered=[{offered_by or '—'}]  country=[{host_country or '—'}]  dl=[{deadline_raw or '—'}]  web=[{'✓' if official_url else '✗'}]")

    if deadline_raw and deadline_passed(deadline_raw):
        print(f"  ⏭  SKIP — deadline passed [{deadline_raw}]"); return None

    parsed_d = parse_deadline(deadline_raw); deadline_out = parsed_d.isoformat() if parsed_d else ""

    # Fallback: visit official link for uni name + deadline
    if official_url and (not offered_by or not deadline_out):
        try:
            print(f"     visiting official page ...")
            await page.goto(official_url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
            await page.wait_for_timeout(2000)
            official_text = clean(await page.evaluate(r"() => { const c=document.body.cloneNode(true); c.querySelectorAll('script,style,nav,header,footer').forEach(n=>n.remove()); return (c.innerText||'').trim().slice(0,5000); }"))
            if not offered_by:
                raw_title = clean(await page.evaluate(r"() => { const h=document.querySelector('h1'); return h?(h.innerText||'').trim():document.title||''; }"))
                NOISE = {'index','home','welcome','apply','apply now','login','sign in','register','application','announcements','news','scholarships','admissions','portal'}
                ct = re.sub(r'\s*[\|\-–—]\s*.+$','',raw_title).strip()
                if 4 < len(ct) < 90 and ct.lower() not in NOISE and not re.match(r'^\d',ct):
                    inst = ['university','institute','college','school','ministry','government','foundation','council','hochschule','universidad']
                    if any(w in ct.lower() for w in inst):
                        offered_by = ct; print(f"     offered_by: {offered_by}")
                    else:
                        sn = clean(await page.evaluate(r"() => { const m=document.querySelector('meta[property=\"og:site_name\"]'); return m?(m.getAttribute('content')||''):''; }"))
                        if sn and 4 < len(sn) < 80: offered_by = sn; print(f"     offered_by (og): {offered_by}")
            if not deadline_out and official_text:
                for pat in [r'(?:deadline|last date|closing date|apply by|due date)\s*[:\-–]\s*(\d{1,2}\s+[A-Za-z]+\s+\d{4})',r'Deadline[^\n]{0,30}?(\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4})']:
                    m = re.search(pat, official_text, re.IGNORECASE)
                    if m:
                        c = m.group(1).strip().rstrip('.'); p = parse_deadline(c)
                        if p:
                            if deadline_passed(c): print(f"  ⏭  SKIP — official deadline passed [{c}]"); return None
                            deadline_out = p.isoformat(); print(f"     deadline from official: {deadline_out}"); break
            await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
            await page.wait_for_timeout(1000)
        except Exception as e:
            print(f"     official page failed: {e}")
            try: await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
            except Exception: pass

    ai_data = await ai_extract(name, fields, text)
    final_degree  = clean_degree(degree_raw  or clean(str(ai_data.get("degree_type",""))))
    final_type    = coverage                 or clean(str(ai_data.get("scholarship_type","")))
    final_country = host_country             or clean(str(ai_data.get("host_country","")))
    final_field   = clean(str(ai_data.get("field_of_study","")))
    if not deadline_out:
        ai_dl = clean(str(ai_data.get("deadline","")))
        if ai_dl:
            p = parse_deadline(ai_dl)
            if p and not deadline_passed(ai_dl): deadline_out = p.isoformat()

    print(f"  ✅  dl=[{deadline_out or 'blank'}]  type=[{final_type or '—'}]  degree=[{final_degree or '—'}]")
    return {
        "Scholarship Name": name, "Scholarship Link": link, "Offered By": offered_by,
        "Official Website": official_url, "Scholarship Type": final_type, "Degree Type": final_degree,
        "Host Country": final_country, "Eligible Nations": nationality,
        "Deadline": deadline_out, "Field of Study": final_field,
    }


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═"*66)
    print("  OpportunitiesCorners.com Scraper  →  Google Sheets")
    print(f"  Pages  : {START_PAGE} → {MAX_PAGES}  |  Tab: {SHEET_TAB}")
    print("═"*66)
    print("\n  🔗  Connecting to Google Sheets ...")
    existing = load_existing_links()
    print(f"\n  🗃  {len(existing)} already-saved links loaded\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True, args=["--no-sandbox","--disable-setuid-sandbox","--disable-dev-shm-usage","--disable-gpu"])
        ctx = await browser.new_context(viewport={"width":1280,"height":900}, extra_http_headers=REQUEST_HEADERS)
        async def block_junk(route):
            if route.request.resource_type in {"image","font","media"}: await route.abort()
            else: await route.continue_()
        listing_page = await ctx.new_page(); detail_page = await ctx.new_page()
        await listing_page.route("**/*", block_junk); await detail_page.route("**/*", block_junk)
        listing_page.set_default_navigation_timeout(NAV_TIMEOUT); detail_page.set_default_navigation_timeout(NAV_TIMEOUT)
        saved = skipped = 0

        for page_num in range(START_PAGE, MAX_PAGES + 1):
            url = BASE_URL if page_num == 1 else PAGE_URL_TMPL.format(n=page_num)
            print(f"\n{'═'*66}\n  📄  PAGE {page_num}/{MAX_PAGES}  →  {url}\n{'═'*66}")
            try:
                await listing_page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                await listing_page.wait_for_timeout(PAGE_WAIT)
            except Exception as e:
                print(f"  ⚠  Page {page_num} failed: {e}"); continue

            cards = await listing_page.evaluate(COLLECT_CARDS_JS)
            if not cards: print(f"  ⚠  No cards — stopping."); break
            new_cards = [c for c in cards if c["link"] not in existing]
            print(f"  {len(cards)} cards  |  {len(cards)-len(new_cards)} saved  |  {len(new_cards)} to scrape\n")

            for i, card in enumerate(new_cards, start=1):
                existing.add(card["link"])
                result = await scrape_one(detail_page, card, seq=i, total=len(new_cards))
                if result: save_one_to_sheets(result); saved += 1; print(f"  💾  Saved  [total={saved}]")
                else: skipped += 1
                await asyncio.sleep(0.5)

        await browser.close()

    print(f"\n{'═'*66}\n  ✅  DONE  |  Saved: {saved}  |  Skipped: {skipped}")
    print(f"  Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}\n{'═'*66}")


if __name__ == "__main__":
    asyncio.run(main())