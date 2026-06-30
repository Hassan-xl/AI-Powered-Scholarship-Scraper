import os
import re
import json
import time
import asyncio
from datetime import datetime, date

from playwright.async_api import async_playwright
from openai import AsyncOpenAI
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

# ==================================================
# CONFIG
# ==================================================
START_URL = "https://www.topuniversities.com/scholarships/scholarships-for-students"
TOTAL_PAGES = 24

FLUSH_EVERY = 5
NAV_TIMEOUT_MS = 60_000
TODAY = date.today()

MAX_CONCURRENT_SCRAPES = 5
MAX_CONCURRENT_OPENAI = 10

# ==================================================
# GOOGLE SHEETS CONFIG
# ==================================================
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "topuniversities_scholarships"

COLUMN_HEADERS = [
    "Scholarship Name", "Scholarship Link", "Official Website",
    "Scholarship Type", "Deadline", "Host Country",
    "Degree Type", "Field of Study",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==================================================
# OPENAI
# ==================================================
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
client = AsyncOpenAI(api_key=OPENAI_API_KEY)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122 Safari/537.36"
    )
}

SCHEMA_EMPTY = {
    "scholarship_type": "",
    "deadline": "",
    "host_country": "",
    "degree_type": "",
    "field_of_study": "",
}

# ==================================================
# GOOGLE SHEETS HELPERS
# ==================================================
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

def save_batch(rows_buffer, retries: int = 3):
    global _header_written
    if not rows_buffer:
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
                for record in rows_buffer
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
            print(f"  ⚠  save_batch error: {e}")
            time.sleep(5)

# ==================================================
# JAVASCRIPT EXTRACTORS
# ==================================================
EXTRACT_MAIN_TEXT_JS = r"""
() => {
  const candidates = [
    'article .field--name-body',
    'article .field--type-text-with-summary',
    'article .node__content',
    'main article',
    'article',
    'main'
  ];
  const strip = (root) => {
    if (!root) return;
    root.querySelectorAll(
      'nav, header, footer, form, .breadcrumb, .pagination, .share, .social'
    ).forEach(n => n.remove());
  };
  for (const sel of candidates) {
    const el = document.querySelector(sel);
    if (el) {
      const clone = el.cloneNode(true);
      strip(clone);
      const text = (clone.innerText || '').trim();
      if (text.length > 200) return text;
    }
  }
  const blocks = Array.from(
    document.querySelectorAll('main div, main section, article div, article section')
  )
    .map(el => ({ el, len: (el.innerText || '').trim().length }))
    .filter(x => x.len > 300)
    .sort((a, b) => b.len - a.len);
  if (blocks[0]) return (blocks[0].el.innerText || '').trim();
  return (document.body?.innerText || '').trim();
}
"""

EXTRACT_CARDS_JS = r"""
() => {
    const cards = document.querySelectorAll('div.scholarship-cards div.scholarship-card');
    const results = [];
    cards.forEach(card => {
        let title = '';
        const titleEl = card.querySelector('div.scholarship-title');
        if (titleEl) title = (titleEl.innerText || '').trim();
        if (!title) {
            const h3 = card.querySelector('h3');
            if (h3) title = (h3.innerText || '').trim();
        }
        if (!title) {
            const h2 = card.querySelector('h2');
            if (h2) title = (h2.innerText || '').trim();
        }

        let link = '';
        const anchor = card.querySelector('div.cta-sec a.yellow-anchor[href]');
        if (anchor) link = anchor.getAttribute('href') || '';
        if (!link) {
            const fallback = card.querySelector('a[href*="/scholarships/"]');
            if (fallback) link = fallback.getAttribute('href') || '';
        }

        if (title && link) results.push({ title, link });
    });
    return results;
}
"""

# ==================================================
# HELPERS
# ==================================================
def fix_encoding(s: str) -> str:
    if not s:
        return s
    try:
        fixed = s.encode("cp1252").decode("utf-8")
        s = fixed
    except (UnicodeEncodeError, UnicodeDecodeError):
        pass
    replacements = {
        "\u2018": "'", "\u2019": "'", "\u201C": '"', "\u201D": '"',
        "\u2013": "-", "\u2014": "-", "\u2026": "...", "\u00A0": " ",
        "\u00E2\u0080\u0099": "'",
    }
    for old, new in replacements.items():
        s = s.replace(old, new)
    s = s.encode("ascii", errors="ignore").decode("ascii")
    return s.strip()

def clean(s: str) -> str:
    s = fix_encoding(s)
    return re.sub(r"\s+", " ", (s or "")).strip()

def clean_degree_type(s: str) -> str:
    if not s:
        return s
    s = fix_encoding(s)
    parts = [p.strip() for p in re.split(r"[,;]+", s) if p.strip()]
    cleaned = []
    for part in parts:
        p = part.strip()
        p = re.sub(r"^Master'?s?$", "Masters", p, flags=re.IGNORECASE)
        p = re.sub(r"^Bachelor'?s?$", "Bachelors", p, flags=re.IGNORECASE)
        if p and p not in cleaned:
            cleaned.append(p)
    return ", ".join(cleaned)

def parse_deadline(deadline_str: str) -> date | None:
    if not deadline_str:
        return None
    deadline_str = clean(deadline_str)
    formats = [
        "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%y",
        "%d-%B-%Y", "%d %B, %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(deadline_str, fmt).date()
        except ValueError:
            continue
    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})",
        deadline_str, re.IGNORECASE,
    )
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %B %Y"
            ).date()
        except ValueError:
            pass
    return None

def is_deadline_passed(deadline_str: str) -> bool:
    d = parse_deadline(deadline_str)
    if d is None:
        return False
    return d < TODAY

def split_degree_types(degree_str: str) -> list:
    if not degree_str or not degree_str.strip():
        return [""]
    parts = [p.strip() for p in re.split(r"[,;]+", fix_encoding(degree_str)) if p.strip()]
    seen, out = set(), []
    for p in parts:
        if p not in seen:
            seen.add(p)
            out.append(p)
    return out if out else [""]

async def wait_listing_ready(page):
    await page.wait_for_selector(
        "div.scholarship-cards div.scholarship-card", timeout=60_000
    )

async def discover_scholarships(page, total_pages=TOTAL_PAGES):
    scholarships = []
    seen = set()
    await page.goto(START_URL, wait_until="domcontentloaded", timeout=60_000)
    await wait_listing_ready(page)

    for page_idx in range(1, total_pages + 1):
        print(f"\n📄 Listing page {page_idx}/{total_pages}")
        await page.evaluate("""
            () => {
                const section = document.querySelector('div.scholarship-cards');
                if (section) section.scrollIntoView({ behavior: 'instant', block: 'end' });
            }
        """)
        await page.wait_for_timeout(300)
        await wait_listing_ready(page)

        card_data = await page.evaluate(EXTRACT_CARDS_JS)
        print(f"  Cards found: {len(card_data)}")

        for item in card_data:
            title = clean(item["title"])
            link  = item["link"].strip()
            if not link:
                continue
            if not link.startswith("http"):
                link = "https://www.topuniversities.com" + link
            if link in seen:
                continue
            seen.add(link)
            scholarships.append({"name": title, "link": link})
            print(f"  ✔ {title}")

        if page_idx < total_pages:
            next_selectors = [
                "ul.pagination a[rel='next']",
                "ul.pagination li:last-child a",
                "nav ul.pagination li:last-child a",
                "a[aria-label='Next']",
            ]
            clicked = False
            for sel in next_selectors:
                try:
                    loc = page.locator(sel).first
                    if await loc.count() == 0:
                        continue
                    await loc.scroll_into_view_if_needed(timeout=3000)
                    await loc.click(timeout=5000)
                    await page.wait_for_timeout(1500)
                    try:
                        await wait_listing_ready(page)
                    except Exception:
                        await page.wait_for_timeout(2000)
                        await wait_listing_ready(page)
                    clicked = True
                    break
                except Exception:
                    continue

            if not clicked:
                try:
                    next_num = page_idx + 1
                    pager = page.locator(
                        f"ul.pagination li a:has-text('{next_num}')"
                    ).first
                    await pager.scroll_into_view_if_needed(timeout=3000)
                    await pager.click(timeout=5000)
                    await page.wait_for_timeout(1500)
                    await wait_listing_ready(page)
                    clicked = True
                except Exception:
                    pass

            if not clicked:
                print(f"  ⚠ Could not navigate to page {page_idx + 1}, stopping discovery.")
                break

    return scholarships

async def extract_official_website(page) -> str:
    for text_match in ("View Scholarship", "Official Website", "Apply", "Visit website"):
        try:
            href = await (
                page.locator(f"a:has-text('{text_match}')").first.get_attribute("href")
            )
            if href and href.startswith("http"):
                return href
        except Exception:
            pass
    try:
        hrefs = await page.eval_on_selector_all(
            "article a[href^='http'], main a[href^='http']",
            "els => els.map(e => e.getAttribute('href'))",
        )
        for h in hrefs or []:
            if h and "topuniversities.com" not in h:
                return h
    except Exception:
        pass
    return ""

def build_prompt(text: str, page_url: str, page_title: str) -> str:
    return (
        "You are a data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON, no explanations, no markdown, no extra text.\n"
        '- If a field is truly not determinable, return "" for it.\n'
        "- Use plain ASCII only. Replace curly quotes with straight quotes. "
        "Write Master's as Masters, Bachelor's as Bachelors.\n"
        "- For degree_type: use only these values separated by commas: "
        "Bachelors, Masters, MBA, Executive MBA, PhD. No apostrophes.\n\n"
        "IMPORTANT for host_country:\n"
        "- The host country is WHERE the university/institution is physically located.\n"
        "- Infer it from the university name, URL, or any contextual clues.\n"
        "  Examples: Clark University -> United States, "
        "Heriot-Watt University -> United Kingdom, "
        "University of Melbourne -> Australia, "
        "ETH Zurich -> Switzerland.\n"
        "- If the page mentions a specific campus country, use that.\n"
        "- If multiple countries, list the primary one.\n\n"
        "JSON schema:\n"
        "{\n"
        '  "scholarship_type": "",\n'
        '  "deadline": "",\n'
        '  "host_country": "",\n'
        '  "degree_type": "",\n'
        '  "field_of_study": ""\n'
        "}\n\n"
        f"Page URL: {page_url}\n"
        f"Page Title: {page_title}\n\n"
        f"Text:\n{text}"
    )

def parse_json_safely(raw: str) -> dict:
    raw = (raw or "").strip()
    try:
        data = json.loads(raw)
        return data if isinstance(data, dict) else dict(SCHEMA_EMPTY)
    except Exception:
        m = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if m:
            try:
                data = json.loads(m.group(0))
                return data if isinstance(data, dict) else dict(SCHEMA_EMPTY)
            except Exception:
                pass
        return dict(SCHEMA_EMPTY)

async def openai_extract(text: str, page_url: str, page_title: str) -> dict:
    prompt = build_prompt(text, page_url, page_title)
    response = await client.chat.completions.create(
        model="gpt-4o-mini",
        messages=[{"role": "user", "content": prompt}],
        temperature=0,
    )
    raw  = (response.choices[0].message.content or "").strip()
    data = parse_json_safely(raw)
    out  = dict(SCHEMA_EMPTY)
    for k in out:
        if k in data and data[k] is not None:
            out[k] = str(data[k]).strip()
    return out

async def scrape_one(
    scholarship: dict,
    browser,
    scrape_sem: asyncio.Semaphore,
    openai_sem: asyncio.Semaphore,
    idx: int,
    total: int,
) -> list | None:
    async with scrape_sem:
        context = await browser.new_context(extra_http_headers=HEADERS)
        await context.route(
            "**/*",
            lambda route: (
                asyncio.ensure_future(route.abort())
                if route.request.resource_type in ("image", "font", "media", "stylesheet")
                else asyncio.ensure_future(route.continue_())
            ),
        )
        page = await context.new_page()
        page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

        name, link = scholarship["name"], scholarship["link"]
        print(f"🔍 ({idx}/{total}) {name}")

        # --- ADDED RETRY LOGIC FOR PAGE LOAD ---
        page_loaded = False
        max_retries = 3
        
        for attempt in range(max_retries):
            try:
                await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
                page_loaded = True
                break
            except Exception as e:
                print(f"  ⚠ Page load failed (attempt {attempt + 1}/{max_retries}): {e}")
                if attempt < max_retries - 1:
                    await asyncio.sleep(2) # brief pause before retrying
        
        if not page_loaded:
            print(f"  ❌ Skipping {name} after {max_retries} failed load attempts.")
            await context.close()
            return None

        text = clean(await page.evaluate(EXTRACT_MAIN_TEXT_JS))
        if len(text) > 9000:
            text = text[:9000]

        official   = await extract_official_website(page)
        page_title = clean(await page.title())
        await context.close()

    async with openai_sem:
        try:
            data = await openai_extract(text, link, page_title)
        except Exception as e:
            print(f"  ⚠ OpenAI failed: {e}")
            data = dict(SCHEMA_EMPTY)

    deadline_raw     = clean(data.get("deadline", ""))
    degree_type_raw  = clean_degree_type(data.get("degree_type", ""))
    host_country     = clean(data.get("host_country", ""))
    scholarship_type = clean(data.get("scholarship_type", ""))
    field_of_study   = clean(data.get("field_of_study", ""))

    deadline_formatted = deadline_raw
    if deadline_raw:
        parsed_date = parse_deadline(deadline_raw)
        if parsed_date:
            if parsed_date < TODAY:
                print(f"  ⏭ SKIPPED (deadline passed: {deadline_raw})")
                return None
            deadline_formatted = parsed_date.strftime("%Y-%m-%d")

    base = {
        "Scholarship Name": clean(name),
        "Scholarship Link": link,
        "Official Website": official,
        "Scholarship Type": scholarship_type,
        "Deadline":         deadline_formatted,
        "Host Country":     host_country,
        "Field of Study":   field_of_study,
    }

    degree_types = split_degree_types(degree_type_raw)
    if len(degree_types) > 1:
        print(f"  ↳ Expanding to {len(degree_types)} rows: {degree_types}")

    rows = []
    for dt in degree_types:
        row = dict(base)
        row["Degree Type"] = dt
        rows.append(row)

    return rows

async def main():
    print("🔗 Connecting to Google Sheets...")
    existing_links = load_existing_links()
    print(f"🗃️ Loaded {len(existing_links)} already-scraped links from Google Sheets (tab: {SHEET_TAB})")

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)

        ctx = await browser.new_context(extra_http_headers=HEADERS)
        await ctx.route(
            "**/*",
            lambda route: (
                asyncio.ensure_future(route.abort())
                if route.request.resource_type in ("image", "font", "media", "stylesheet")
                else asyncio.ensure_future(route.continue_())
            ),
        )
        listing_page = await ctx.new_page()
        listing_page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

        print("🌐 Discovering scholarship links...")
        scholarships = await discover_scholarships(listing_page, TOTAL_PAGES)
        print(f"\n📦 Total unique scholarships found: {len(scholarships)}")
        await ctx.close()

        scholarships = [
            s for s in scholarships
            if s["link"].strip() not in existing_links
        ]
        print(f"✅ Proceeding with {len(scholarships)} new scholarships to scrape (skipped {len(existing_links)} already scraped)\n")

        print(f"\n🧭 Scraping detail pages (batch size={FLUSH_EVERY})...\n")
        scrape_sem = asyncio.Semaphore(MAX_CONCURRENT_SCRAPES)
        openai_sem = asyncio.Semaphore(MAX_CONCURRENT_OPENAI)

        saved   = 0
        skipped = 0
        total   = len(scholarships)

        for batch_start in range(0, total, FLUSH_EVERY):
            batch_end = min(batch_start + FLUSH_EVERY, total)
            batch     = scholarships[batch_start:batch_end]

            tasks = [
                scrape_one(s, browser, scrape_sem, openai_sem, batch_start + j + 1, total)
                for j, s in enumerate(batch)
            ]

            batch_results = await asyncio.gather(*tasks, return_exceptions=True)

            valid_records = []
            for r in batch_results:
                if isinstance(r, list):
                    valid_records.extend(r)
                elif r is None:
                    skipped += 1
                elif isinstance(r, Exception):
                    print(f"  ⚠ Task exception: {r}")
                    skipped += 1

            if valid_records:
                save_batch(valid_records)
                saved += len(valid_records)

            print(
                f"💾 Batch [{batch_start+1}-{batch_end}] done — "
                f"{len(valid_records)} rows saved (Total rows: {saved}), "
                f"{skipped} skipped/failed → Google Sheets\n"
            )

        await browser.close()

    print(f"\n✅ DONE — {saved} rows saved, {skipped} skipped (expired/failed)")
    print(f"🔗 Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}")

if __name__ == "__main__":
    asyncio.run(main())