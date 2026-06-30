"""
IDP Find-a-Scholarship Scraper  →  Google Sheets Version
==========================================================
Site   : https://www.idp.com/find-a-scholarship/?page=1
Output : Google Sheets → Tab: idp_scholarships

Structure
---------
- ~6300 scholarships, 12 per page, ~525 pages
- Each card → visit uni page → grab official website
- Each card → visit detail page → AI extract fields
- Pagination via URL: ?page=1, ?page=2 ...
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
BASE_URL          = "https://www.idp.com/find-a-scholarship/?page={n}"
START_PAGE        = 1
MAX_PAGES         = 525
FLUSH_EVERY       = 12
NAV_TIMEOUT       = 60_000
PAGE_WAIT         = 2000
DETAIL_WAIT       = 2000
MAX_AI_CONCURRENT = 8
TODAY             = date.today()

BLOCKED_TYPES = {"image", "font", "media"}

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
SHEET_TAB        = "idp_scholarships"

COLUMN_HEADERS = [
    "Scholarship Name", "Scholarship Link", "Official Website",
    "Scholarship Type", "Deadline", "Host Country",
    "Degree Type", "Field of Study",
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
            # Write header IMMEDIATELY on tab creation
            _gs_worksheet.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
            _header_written = True
            print(f"  📋  Created tab '{SHEET_TAB}' with headers")
    return _gs_worksheet


def load_existing_links() -> set:
    """Load already-saved Scholarship Links for duplicate check."""
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


def flush_to_sheets(records: list[dict], retries: int = 3):
    """Append a batch of records to Google Sheets."""
    global _header_written
    if not records:
        return
    for attempt in range(retries):
        try:
            ws = _get_worksheet()
            if not _header_written:
                ws.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
                _header_written = True
                time.sleep(1)
            rows = [
                [str(record.get(col, "") or "") for col in COLUMN_HEADERS]
                for record in records
            ]
            ws.append_rows(rows, value_input_option="USER_ENTERED")
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
            print(f"  ⚠  flush_to_sheets error: {e}")
            time.sleep(5)


# ══════════════════════════════════════════════════════════════
# JAVASCRIPT — EXTRACT CARDS FROM LISTING PAGE
# ══════════════════════════════════════════════════════════════
EXTRACT_CARDS_JS = r"""
() => {
    const results = [];
    const cards = document.querySelectorAll('div.h-full.interactive-card');
    cards.forEach(card => {
        let name = '', link = '';
        const titleA = card.querySelector('a.h4, a[class*="font-heading"], a[class*="h4"]');
        if (titleA) {
            name = (titleA.innerText || '').trim();
            link = titleA.getAttribute('href') || '';
            if (link && !link.startsWith('http')) link = 'https://www.idp.com' + link;
        }

        let university = '', uniPageLink = '';
        const allCardAnchors = card.querySelectorAll('a');
        for (const a of allCardAnchors) {
            const href = (a.getAttribute('href') || '');
            if (href.includes('/universities-and-colleges/')) {
                uniPageLink = href.startsWith('http') ? href : 'https://www.idp.com' + href;
                university  = (a.innerText || '').trim();
                break;
            }
        }
        if (!university) {
            const uniP = card.querySelector('p.text-small, p[class*="truncate"]');
            if (uniP) university = (uniP.innerText || '').trim();
        }

        let country = '', degree = '', funding = '', deadline = '';
        const liTexts = Array.from(card.querySelectorAll('ul li'))
            .map(li => (li.innerText || '').trim()).filter(Boolean);
        liTexts.forEach(txt => {
            const low = txt.toLowerCase();
            if (low.includes('deadline:'))
                deadline = txt.replace(/^deadline[:\s]*/i, '').trim();
            else if (low.includes('funding type:'))
                funding  = txt.replace(/^funding type[:\s]*/i, '').trim();
            else if (low.includes('undergraduate') || low.includes('postgraduate') ||
                     low.includes('bachelor') || low.includes('master') ||
                     low.includes('phd') || low.includes('doctorate'))
                degree = txt;
            else if (!country && !low.includes('value') && !low.includes('award'))
                country = txt;
        });

        if (name && link) {
            results.push({ name, link, university, uniPageLink,
                           country, degree, funding, deadline });
        }
    });
    return results;
}
"""

EXTRACT_UNI_WEBSITE_JS = r"""
() => {
    for (const a of document.querySelectorAll('a')) {
        const txt  = (a.innerText || '').trim().toLowerCase();
        const href = (a.getAttribute('href') || '').trim();
        if (
            href.startsWith('http') &&
            !href.includes('idp.com') &&
            (txt.includes('visit the university') ||
             txt.includes('university website') ||
             txt === 'website')
        ) { return href; }
    }
    return '';
}
"""

EXTRACT_DETAIL_JS = r"""
() => {
    const selectors = [
        '[class*="scholarship-detail"]',
        '[class*="scholarship-content"]',
        'main', 'article', '#root',
    ];
    let bestText = '';
    for (const sel of selectors) {
        const el = document.querySelector(sel);
        if (el) {
            const clone = el.cloneNode(true);
            clone.querySelectorAll(
                'script,style,nav,header,footer,button,[class*="widget"],[class*="popup"]'
            ).forEach(n => n.remove());
            const t = (clone.innerText || '').trim();
            if (t.length > bestText.length) bestText = t;
        }
    }
    if (!bestText) bestText = (document.body?.innerText || '').trim();
    return bestText;
}
"""

# ══════════════════════════════════════════════════════════════
# HELPERS
# ══════════════════════════════════════════════════════════════

def safe_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, (list, tuple)):
        return ", ".join(str(v).strip() for v in val if v)
    return str(val).strip()


def clean(s) -> str:
    s = safe_str(s)
    for old, new in {
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
        "\u2013": "-", "\u2014": "-", "\u2026": "...", "\u00a0": " ",
    }.items():
        s = s.replace(old, new)
    return re.sub(r"\s+", " ", s).strip()


def clean_degree(s: str) -> str:
    parts = [p.strip() for p in re.split(r"[,;/]+", s) if p.strip()]
    out, seen = [], set()
    for p in parts:
        p = re.sub(r"^Master'?s?$",   "Masters",   p, flags=re.IGNORECASE)
        p = re.sub(r"^Bachelor'?s?$", "Bachelors", p, flags=re.IGNORECASE)
        if p.lower() not in seen:
            seen.add(p.lower())
            out.append(p)
    return ", ".join(out)


def parse_deadline(s: str) -> date | None:
    s = clean(s)
    if not s or s.lower() in ("none", "n/a", "-", ""):
        return None
    formats = [
        "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
        "%d-%B-%Y", "%d %B, %Y", "%d %b, %Y",
        "%B %Y", "%b %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})", s, re.I)
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)} {m.group(2).title()} {m.group(3)}", "%d %B %Y"
            ).date()
        except ValueError:
            pass
    m = re.search(
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})", s, re.I)
    if m:
        try:
            return datetime.strptime(
                f"{m.group(2)} {m.group(1).title()} {m.group(3)}", "%d %B %Y"
            ).date()
        except ValueError:
            pass
    return None


def deadline_passed(s: str) -> bool:
    d = parse_deadline(s)
    return d is not None and d < TODAY


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(text: str, name: str, card: dict) -> str:
    return (
        "You are a precise data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON — no markdown, no commentary.\n"
        "- All values must be plain ASCII strings (never arrays or null).\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated from: Bachelors, Masters, MBA, PhD, Executive MBA.\n"
        "- scholarship_type: e.g. 'Fully Funded', 'Partial', 'Fee waiver', 'Tuition discount'.\n"
        "- deadline: exact date string as found. Return '' if not mentioned.\n"
        "- host_country: country where student will study.\n"
        "- field_of_study: comma-separated subject fields.\n\n"
        "JSON schema:\n"
        "{\n"
        '  "scholarship_type": "",\n'
        '  "deadline": "",\n'
        '  "host_country": "",\n'
        '  "degree_type": "",\n'
        '  "field_of_study": ""\n'
        "}\n\n"
        f"Scholarship Name : {name}\n"
        f"University       : {card.get('university','')}\n"
        f"Country (card)   : {card.get('country','')}\n"
        f"Degree (card)    : {card.get('degree','')}\n"
        f"Funding (card)   : {card.get('funding','')}\n"
        f"Deadline (card)  : {card.get('deadline','')}\n\n"
        f"Detail Page Text:\n{text[:7000]}"
    )


def parse_json_safe(raw: str) -> dict:
    raw = re.sub(r"^```[a-z]*\s*", "", raw.strip(), flags=re.I)
    raw = re.sub(r"\s*```$", "", raw).strip()
    try:
        d = json.loads(raw)
        return d if isinstance(d, dict) else dict(EMPTY)
    except Exception:
        m = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if m:
            try:
                d = json.loads(m.group(0))
                return d if isinstance(d, dict) else dict(EMPTY)
            except Exception:
                pass
    return dict(EMPTY)


async def ai_extract(text: str, name: str, card: dict,
                     sem: asyncio.Semaphore) -> dict:
    async with sem:
        try:
            resp = await ai_client.chat.completions.create(
                model="gpt-4o-mini",
                temperature=0,
                messages=[
                    {"role": "system",
                     "content": "Structured data extraction. Return only valid JSON."},
                    {"role": "user", "content": build_prompt(text, name, card)},
                ],
                max_tokens=400,
            )
            raw = resp.choices[0].message.content or ""
        except Exception as e:
            print(f"  ⚠  OpenAI [{name[:30]}]: {e}")
            return dict(EMPTY)
    data = parse_json_safe(raw)
    out  = dict(EMPTY)
    for k in out:
        if k in data and data[k] is not None:
            out[k] = clean(str(data[k]))
    return out


# ══════════════════════════════════════════════════════════════
# SCRAPE ONE SCHOLARSHIP
# ══════════════════════════════════════════════════════════════

async def scrape_one(
    page, card: dict, seq: int, total: int, ai_sem: asyncio.Semaphore,
) -> dict | None:

    name         = clean(card["name"])
    link         = card["link"]
    uni          = clean(card.get("university",   ""))
    uni_page     = clean(card.get("uniPageLink",  ""))
    country      = clean(card.get("country",      ""))
    degree       = clean(card.get("degree",       ""))
    funding      = clean(card.get("funding",      ""))
    deadline_raw = clean(card.get("deadline",     ""))

    print(f"\n  🔎  ({seq}/{total}) {name[:50]}")
    print(f"       uni page  : {uni_page[:70] or 'NOT FOUND'}")
    print(f"       deadline  : {deadline_raw or 'none on card'}")

    if deadline_raw and deadline_passed(deadline_raw):
        print(f"  ⏭  SKIP expired: {deadline_raw}")
        return None

    parsed       = parse_deadline(deadline_raw)
    deadline_out = parsed.isoformat() if parsed else ""

    # Visit IDP university page → official website
    official_url = ""
    if uni_page:
        try:
            await page.goto(uni_page, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
            await page.wait_for_timeout(2000)
            official_url = clean(await page.evaluate(EXTRACT_UNI_WEBSITE_JS))
            print(f"       official  : {official_url or 'NOT FOUND'}")
        except Exception as e:
            print(f"       uni page error: {e}")
    else:
        print(f"       ⚠ No uni page link on card!")

    # Visit scholarship detail page → AI text
    try:
        await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await page.wait_for_timeout(DETAIL_WAIT)
    except Exception as e:
        print(f"  ⚠  Detail load failed: {e}")
        return None

    text     = clean(await page.evaluate(EXTRACT_DETAIL_JS))
    ai_data  = await ai_extract(text, name, card, ai_sem)

    degree_out   = clean_degree(degree or clean(ai_data.get("degree_type",      "")))
    host_country = country               or clean(ai_data.get("host_country",    ""))
    schol_type   = funding               or clean(ai_data.get("scholarship_type",""))

    print(f"  ✅  uni=[{uni[:25]}]  dl=[{deadline_out or 'blank'}]  web=[{official_url[:35] or 'none'}]")

    return {
        "Scholarship Name": name,
        "Scholarship Link": link,
        "Official Website": official_url,
        "Scholarship Type": schol_type,
        "Deadline":         deadline_out,
        "Host Country":     host_country,
        "Degree Type":      degree_out,
        "Field of Study":   clean(ai_data.get("field_of_study", "")),
    }


# ══════════════════════════════════════════════════════════════
# PROCESS ONE LISTING PAGE
# ══════════════════════════════════════════════════════════════

async def process_listing_page(
    listing_page, detail_page, page_num: int,
    existing: set, ai_sem: asyncio.Semaphore,
) -> list[dict]:

    url = BASE_URL.format(n=page_num)
    try:
        await listing_page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await listing_page.wait_for_timeout(PAGE_WAIT)
    except Exception as e:
        print(f"\n  ⚠  Page {page_num} load failed: {e}")
        return []

    cards = await listing_page.evaluate(EXTRACT_CARDS_JS)
    if not cards:
        print(f"\n  ⚠  Page {page_num} — no cards (end of results?)")
        return []

    new_cards = [c for c in cards if c["link"] not in existing]
    print(f"\n  📄  Page {page_num}  —  {len(cards)} cards  "
          f"({len(cards)-len(new_cards)} saved, {len(new_cards)} new)")

    records = []
    for i, card in enumerate(new_cards, start=1):
        res = await scrape_one(detail_page, card, seq=i,
                               total=len(new_cards), ai_sem=ai_sem)
        if res:
            records.append(res)
            existing.add(card["link"])

    return records


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═" * 64)
    print("  IDP Find-a-Scholarship Scraper  →  Google Sheets")
    print(f"  Pages  : {START_PAGE} → {MAX_PAGES}")
    print(f"  Tab    : {SHEET_TAB}")
    print("═" * 64)

    print("\n  🔗  Connecting to Google Sheets ...")
    existing = load_existing_links()
    print(f"  🗃   {len(existing)} already-saved links loaded\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=["--no-sandbox", "--disable-setuid-sandbox",
                  "--disable-dev-shm-usage", "--disable-gpu"],
        )
        ctx = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            extra_http_headers=REQUEST_HEADERS,
        )

        async def block_junk(route):
            if route.request.resource_type in BLOCKED_TYPES:
                await route.abort()
            else:
                await route.continue_()

        listing_page = await ctx.new_page()
        detail_page  = await ctx.new_page()
        await listing_page.route("**/*", block_junk)
        await detail_page.route("**/*",  block_junk)
        listing_page.set_default_navigation_timeout(NAV_TIMEOUT)
        detail_page.set_default_navigation_timeout(NAV_TIMEOUT)

        ai_sem       = asyncio.Semaphore(MAX_AI_CONCURRENT)
        saved        = 0
        empty_streak = 0
        buf: list[dict] = []

        for page_num in range(START_PAGE, MAX_PAGES + 1):
            records = await process_listing_page(
                listing_page, detail_page, page_num, existing, ai_sem,
            )

            buf.extend(records)
            saved += len(records)

            if buf and (saved % FLUSH_EVERY == 0 or len(buf) >= FLUSH_EVERY):
                flush_to_sheets(buf)
                print(f"\n  💾  Flushed {len(buf)} to Sheets  [total={saved}]\n")
                buf.clear()

            if len(records) == 0 and page_num > START_PAGE:
                empty_streak += 1
                if empty_streak >= 3:
                    print(f"\n  ✅  3 empty pages — done!")
                    break
            else:
                empty_streak = 0

        if buf:
            flush_to_sheets(buf)
            print(f"\n  💾  Final flush — {len(buf)} records")

        await browser.close()

    print("\n" + "═" * 64)
    print("  ✅  DONE")
    print(f"  Saved : {saved}")
    print(f"  Sheet : https://docs.google.com/spreadsheets/d/{SHEET_ID}")
    print("═" * 64)


if __name__ == "__main__":
    asyncio.run(main())