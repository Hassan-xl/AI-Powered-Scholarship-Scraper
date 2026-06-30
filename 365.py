"""
Scholarships365.info Scraper  →  Google Sheets Version
=========================================================
Site   : https://scholarships365.info/masters-scholarships/
Output : Google Sheets → Tab: s365_scholarships

NOTE: Same columns as BrightScholarship/OpportunitiesCorners:
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
BASE_URL      = "https://scholarships365.info/masters-scholarships/page/1/"
PAGE_URL_TMPL = "https://scholarships365.info/masters-scholarships/page/{n}/"
START_PAGE    = 1
MAX_PAGES     = 68
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

BLOCKED_TYPES = {"image", "font", "media"}

# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS CONFIG
# ══════════════════════════════════════════════════════════════
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "s365_scholarships"

# Same columns as BrightScholarship / OpportunitiesCorners
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
    "offered_by":       "",
    "scholarship_type": "",
    "deadline":         "",
    "host_country":     "",
    "degree_type":      "",
    "field_of_study":   "",
    "eligible_nations": "",
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
            # Write header IMMEDIATELY on tab creation
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
    """Save a single record immediately to Google Sheets."""
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
    const selectors = [
        'div.entry-header h2 a[href]', 'div.post h2 a[href]',
        'h2.entry-title a[href]', 'article h2 a[href]', '.entry-title a[href]',
    ];
    for (const sel of selectors) {
        document.querySelectorAll(sel).forEach(a => {
            const href  = (a.getAttribute('href') || '').trim();
            const title = (a.innerText || a.textContent || '').trim();
            if (!href || !title || seen.has(href)) return;
            if (!href.includes('scholarships365.info')) return;
            if (href.includes('/category/') || href.includes('/tag/') ||
                href.includes('/page/') || href.includes('/masters-scholarships/') ||
                href.includes('/author/')) return;
            seen.add(href);
            results.push({ title, link: href });
        });
        if (results.length > 0) break;
    }
    if (results.length === 0) {
        document.querySelectorAll('a[href*="scholarships365.info"]').forEach(a => {
            const href  = (a.getAttribute('href') || '').trim();
            const title = (a.innerText || '').trim();
            if (!href || !title || seen.has(href)) return;
            if (href.includes('/category/') || href.includes('/tag/') ||
                href.includes('/page/') || href.includes('/masters-scholarships/') ||
                href.includes('/author/') || href.endsWith('.info/') ||
                href.includes('#') || title.length < 10) return;
            seen.add(href);
            results.push({ title, link: href });
        });
    }
    return results;
}
"""

EXTRACT_DETAIL_JS = r"""
() => {
    let officialUrl = '';

    // Method A: btn-danger button inside anchor
    document.querySelectorAll('a[target="_blank"]').forEach(a => {
        if (officialUrl) return;
        const btn = a.querySelector('button.btn-danger, button[class*="btn-danger"]');
        if (btn) {
            const href = (a.getAttribute('href') || '').trim();
            if (href.startsWith('http') && !href.includes('scholarships365.info'))
                officialUrl = href;
        }
    });

    // Method B: text = "Official Website"
    if (!officialUrl) {
        document.querySelectorAll('a, button').forEach(el => {
            if (officialUrl) return;
            const txt = (el.innerText || '').trim().toLowerCase();
            if (txt === 'official website' || txt === 'official link') {
                const href = el.tagName === 'A'
                    ? (el.getAttribute('href') || '')
                    : (el.closest('a') ? el.closest('a').getAttribute('href') : '');
                if (href && href.startsWith('http') && !href.includes('scholarships365.info'))
                    officialUrl = href.trim();
            }
        });
    }

    // Method C: Apply Online button
    if (!officialUrl) {
        document.querySelectorAll('a[target="_blank"]').forEach(a => {
            if (officialUrl) return;
            const btn = a.querySelector('button.btn-success, button[class*="btn-success"]');
            if (!btn) return;
            if ((btn.innerText || '').toLowerCase().includes('apply')) {
                const href = (a.getAttribute('href') || '').trim();
                if (href.startsWith('http') && !href.includes('scholarships365.info'))
                    officialUrl = href;
            }
        });
    }

    const content = document.querySelector(
        '#site-content, .site-content, div.post-content, div.entry-content, ' +
        '.details-news, .left-content, article, main'
    );
    const clone = content ? content.cloneNode(true) : document.body.cloneNode(true);
    clone.querySelectorAll(
        'script, style, nav, header, footer, .social-box, .pagination, ' +
        '[class*="adsbygoogle"], [id*="ezoic"], [class*="related"], ' +
        '[class*="comment"], [class*="share"]'
    ).forEach(n => n.remove());
    const text = (clone.innerText || '').trim();

    return { officialUrl, text };
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
    if not s or s.lower() in ("none","n/a","-","open","ongoing","rolling",""): return None
    s = re.sub(r'(\d+)(st|nd|rd|th)\b', r'\1', s, flags=re.IGNORECASE)
    s = re.sub(r'\s+(for|and)\s+.+$', '', s, flags=re.IGNORECASE).strip()
    for fmt in ["%d %B %Y","%d %b %Y","%B %d, %Y","%b %d, %Y","%Y-%m-%d","%d/%m/%Y","%m/%d/%Y","%d-%B-%Y","%d-%b-%Y","%d %B, %Y","%d %b, %Y","%B %Y","%b %Y"]:
        try: return datetime.strptime(s, fmt).date()
        except ValueError: pass
    for pat in [
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)[,\s]+(\d{4})",
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
    ]:
        m = re.search(pat, s, re.I)
        if m:
            try:
                if m.lastindex == 3 and m.group(1).isdigit():
                    return datetime.strptime(f"{m.group(1)} {m.group(2).title()} {m.group(3)}", "%d %B %Y").date()
                elif m.lastindex == 3:
                    return datetime.strptime(f"{m.group(2)} {m.group(1).title()} {m.group(3)}", "%d %B %Y").date()
            except ValueError: pass
    return None

def deadline_passed(s: str) -> bool:
    d = parse_deadline(s)
    return d is not None and d < TODAY

def find_deadline_in_text(text: str) -> str:
    patterns = [
        r'(?:deadline|last date|closing date|apply before|due date|last day|apply by)\s*[:\-–]\s*(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})',
        r'(?:deadline|last date|closing date|apply before|due date|last day|apply by)\s*[:\-–]\s*([A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})',
        r'(?:deadline|last date|closing date|apply before|due date|last day|apply by)\s*[:\-–]\s*(\d{1,2}[\/\-]\d{1,2}[\/\-]\d{4})',
        r'(?:deadline|last date|closing date|apply before|due date|last day|apply by)\s*[:\-–]\s*(\d{4}[\/\-]\d{2}[\/\-]\d{2})',
        r'(?:deadline|last date)[^\n]{0,40}\n\s*(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]+\s+\d{4})',
        r'(?:deadline|last date)[^\n]{0,40}\n\s*([A-Za-z]+\s+\d{1,2}(?:st|nd|rd|th)?,?\s+\d{4})',
        r'Deadline\D{0,20}(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]{3,9}\s+20\d{2})',
        r'Last Date\D{0,20}(\d{1,2}(?:st|nd|rd|th)?\s+[A-Za-z]{3,9}\s+20\d{2})',
        r'(?:deadline|last date|due date)[^\n]{0,50}(20\d{2}-\d{2}-\d{2})',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            candidate = m.group(1).strip().rstrip('.')
            if parse_deadline(candidate):
                return candidate
    return ""


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(name: str, text: str, found_deadline: str, official_url: str) -> str:
    return (
        "You are a precise scholarship data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON — no markdown, no commentary.\n"
        "- All values must be plain ASCII strings (no arrays, no null).\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated from: Bachelors, Masters, MBA, PhD, PostDoc.\n"
        "- scholarship_type: e.g. 'Fully Funded', 'Partial Funding', 'Tuition waiver'.\n"
        "- deadline: scan the ENTIRE text for any date near words like 'deadline', "
        "'last date', 'closing date', 'apply before', 'due date'. "
        "Return the EXACT date string found (e.g. '10 April 2026'). '' if truly not found.\n"
        "- offered_by: university, ministry, or organization offering the scholarship.\n"
        "- host_country: country where the student will STUDY.\n"
        "- eligible_nations: nationalities eligible to apply.\n"
        "- field_of_study: comma-separated subject fields if mentioned.\n\n"
        '{"offered_by":"","scholarship_type":"","deadline":"","host_country":"","degree_type":"","field_of_study":"","eligible_nations":""}\n\n'
        f"Scholarship Name : {name}\n"
        f"Official URL     : {official_url or 'not found'}\n"
        f"Pre-found deadline: {found_deadline or 'none found'}\n\n"
        f"FULL PAGE TEXT:\n{text[:9000]}"
    )


async def ai_extract(name: str, text: str, found_deadline: str, official_url: str) -> dict:
    try:
        resp = await ai_client.chat.completions.create(
            model="gpt-4o-mini", temperature=0,
            messages=[
                {"role": "system", "content": "Scholarship data extraction engine. Return only valid JSON. Hunt carefully through the entire text for deadline dates."},
                {"role": "user",   "content": build_prompt(name, text, found_deadline, official_url)},
            ],
            max_tokens=500,
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

async def scrape_one(page, card: dict, seq: int, total: int) -> dict | None:
    name = clean(card.get("title", ""))
    link = card["link"]

    print(f"\n  ── ({seq}/{total}) {name[:60]}")

    try:
        await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await page.wait_for_timeout(DETAIL_WAIT)
    except Exception as e:
        print(f"     ⚠ Load failed: {e}"); return None

    data         = await page.evaluate(EXTRACT_DETAIL_JS)
    official_url = clean(data.get("officialUrl", ""))
    text         = clean(data.get("text", ""))

    if not text or len(text) < 50:
        print(f"     ⚠ Empty text — skipping"); return None

    found_deadline = find_deadline_in_text(text)
    print(f"     web=[{'✓' if official_url else '✗'}]  pre-deadline=[{found_deadline or '—'}]  text={len(text)}chars")

    if found_deadline and deadline_passed(found_deadline):
        print(f"  ⏭  SKIP — deadline passed [{found_deadline}]"); return None

    ai_data = await ai_extract(name, text, found_deadline, official_url)

    offered_by   = clean(str(ai_data.get("offered_by",       "")))
    schol_type   = clean(str(ai_data.get("scholarship_type", "")))
    host_country = clean(str(ai_data.get("host_country",     "")))
    degree_raw   = clean(str(ai_data.get("degree_type",      "")))
    field_study  = clean(str(ai_data.get("field_of_study",   "")))
    nationality  = clean(str(ai_data.get("eligible_nations", "")))
    dl_from_ai   = clean(str(ai_data.get("deadline",         "")))

    deadline_str = found_deadline or dl_from_ai

    if deadline_str and deadline_passed(deadline_str):
        print(f"  ⏭  SKIP — deadline passed [{deadline_str}]"); return None

    parsed_d     = parse_deadline(deadline_str)
    deadline_out = parsed_d.isoformat() if parsed_d else ""

    print(f"  ✅  offered=[{offered_by[:30] or '—'}]  country=[{host_country or '—'}]  dl=[{deadline_out or 'blank'}]")

    return {
        "Scholarship Name": name,
        "Scholarship Link": link,
        "Offered By":       offered_by,
        "Official Website": official_url,
        "Scholarship Type": schol_type,
        "Degree Type":      clean_degree(degree_raw),
        "Host Country":     host_country,
        "Eligible Nations": nationality,
        "Deadline":         deadline_out,
        "Field of Study":   field_study,
    }


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═" * 66)
    print("  Scholarships365.info Scraper  →  Google Sheets")
    print(f"  Pages  : {START_PAGE} → {MAX_PAGES}")
    print(f"  Tab    : {SHEET_TAB}")
    print("═" * 66)

    print("\n  🔗  Connecting to Google Sheets ...")
    existing = load_existing_links()
    print(f"\n  🗃  {len(existing)} already-saved links loaded\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox","--disable-setuid-sandbox","--disable-dev-shm-usage","--disable-gpu"],
        )
        ctx = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            extra_http_headers=REQUEST_HEADERS,
        )

        async def block_junk(route):
            if route.request.resource_type in BLOCKED_TYPES: await route.abort()
            else: await route.continue_()

        listing_page = await ctx.new_page()
        detail_page  = await ctx.new_page()
        await listing_page.route("**/*", block_junk)
        await detail_page.route("**/*",  block_junk)
        listing_page.set_default_navigation_timeout(NAV_TIMEOUT)
        detail_page.set_default_navigation_timeout(NAV_TIMEOUT)

        saved   = 0
        skipped = 0

        for page_num in range(START_PAGE, MAX_PAGES + 1):
            url = BASE_URL if page_num == 1 else PAGE_URL_TMPL.format(n=page_num)
            print(f"\n{'═'*66}\n  📄  PAGE {page_num}/{MAX_PAGES}  →  {url}\n{'═'*66}")

            try:
                await listing_page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                await listing_page.wait_for_timeout(PAGE_WAIT)
            except Exception as e:
                print(f"  ⚠  Page {page_num} failed: {e} — skipping"); continue

            cards = await listing_page.evaluate(COLLECT_CARDS_JS)
            if not cards:
                debug = await listing_page.evaluate(r"""
                    () => ({ url: window.location.href, h2: document.querySelectorAll('h2').length,
                             posts: document.querySelectorAll('div.post').length,
                             links: document.querySelectorAll('a[href*="scholarships365"]').length })
                """)
                print(f"  🔍 Debug: {debug}")
                print(f"  ⚠  No cards — stopping."); break

            new_cards = [c for c in cards if c["link"] not in existing]
            print(f"  {len(cards)} cards  |  {len(cards)-len(new_cards)} saved  |  {len(new_cards)} to scrape\n")

            for i, card in enumerate(new_cards, start=1):
                existing.add(card["link"])
                result = await scrape_one(detail_page, card, seq=i, total=len(new_cards))
                if result:
                    save_one_to_sheets(result)
                    saved += 1
                    print(f"  💾  Saved  [total={saved}]")
                else:
                    skipped += 1
                await asyncio.sleep(0.5)

        await browser.close()

    print(f"\n{'═'*66}\n  ✅  DONE  |  Saved: {saved}  |  Skipped: {skipped}")
    print(f"  Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}\n{'═'*66}")


if __name__ == "__main__":
    asyncio.run(main())