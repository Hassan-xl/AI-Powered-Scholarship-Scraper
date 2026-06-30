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

# ==================================================
# CONFIG
# ==================================================
BASE_URL        = "https://scholarshipscorner.website/scholarships/"
PAGE_URL_TMPL   = "https://scholarshipscorner.website/scholarships/page/{n}/"
TOTAL_PAGES     = 32          # ← bumped to 32 to capture all pages
NAV_TIMEOUT_MS  = 60_000
TODAY           = date.today()

MAX_CONCURRENT_BROWSERS = 6
MAX_CONCURRENT_OPENAI   = 10
FLUSH_EVERY             = 5

# ==================================================
# GOOGLE SHEETS CONFIG
# ==================================================
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "scholarships_corner"

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

# ==================================================
# OPENAI
# ==================================================
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
ai_client = AsyncOpenAI(api_key=OPENAI_API_KEY)

# ==================================================
# HARDCODED COLUMN HEADERS  ← do not rename until testing done
# ==================================================
COLUMN_HEADERS = [
    "Scholarship Name",
    "Scholarship Link",
    "Official Website",
    "Scholarship Type",
    "Deadline",
    "Host Country",
    "Degree Type",
    "Field of Study",
]

# ==================================================
# HEADERS / BLOCKED TYPES
# ==================================================
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "en-US,en;q=0.9",
}

BLOCKED_TYPES = {"image", "font", "media", "stylesheet"}

EMPTY_SCHEMA = {
    "scholarship_type": "",
    "deadline":         "",
    "host_country":     "",
    "degree_type":      "",
    "field_of_study":   "",
}

# ==================================================
# DEGREE NORMALISATION MAP
# Every variant we've seen → canonical label
# ==================================================
DEGREE_NORM = {
    # Bachelors
    "bachelor":         "Bachelors",
    "bachelors":        "Bachelors",
    "bachelor's":       "Bachelors",
    "undergraduate":    "Bachelors",
    "ug":               "Bachelors",
    "b.sc":             "Bachelors",
    "bsc":              "Bachelors",
    "b.a":              "Bachelors",
    "ba":               "Bachelors",
    "b.eng":            "Bachelors",
    # Masters
    "master":           "Masters",
    "masters":          "Masters",
    "master's":         "Masters",
    "msc":              "Masters",
    "m.sc":             "Masters",
    "ma":               "Masters",
    "m.a":              "Masters",
    "m.eng":            "Masters",
    "postgraduate":     "Masters",
    "pg":               "Masters",
    "graduate":         "Masters",
    # MBA
    "mba":              "MBA",
    # Executive MBA
    "executive mba":    "Executive MBA",
    "emba":             "Executive MBA",
    # PhD
    "phd":              "PhD",
    "ph.d":             "PhD",
    "doctorate":        "PhD",
    "doctoral":         "PhD",
    "dphil":            "PhD",
    # Postdoctoral
    "postdoctoral":     "Postdoctoral",
    "postdoc":          "Postdoctoral",
}

CANONICAL_DEGREE_ORDER = ["Bachelors", "Masters", "MBA", "Executive MBA", "PhD", "Postdoctoral"]


def normalise_degrees(raw: str) -> list[str]:
    """
    Parse a raw degree string and return a deduplicated list of
    canonical degree labels (e.g. ['Bachelors', 'Masters', 'PhD']).
    Unknown tokens are kept as-is (title-cased).
    """
    if not raw:
        return []

    # Split on common separators
    tokens = re.split(r"[,;/|&]+|\band\b|\bor\b|\bwith\b", raw, flags=re.IGNORECASE)
    seen   = {}   # canonical → True (ordered dict style via insertion order)

    for tok in tokens:
        tok = tok.strip().rstrip(".")
        if not tok:
            continue

        # Check full token first, then lowercase
        key = tok.lower()

        # Try exact match
        if key in DEGREE_NORM:
            canon = DEGREE_NORM[key]
        else:
            # Try partial / substring match (longest match wins)
            canon = None
            for pattern, label in DEGREE_NORM.items():
                if pattern in key:
                    if canon is None or len(pattern) > len(
                        next(p for p, l in DEGREE_NORM.items() if l == canon), ""
                    ):
                        canon = label
            if canon is None:
                # Keep unknown but clean it
                canon = tok.title()

        seen[canon] = True

    # Sort by canonical order, unknowns at the end
    def sort_key(d):
        try:
            return CANONICAL_DEGREE_ORDER.index(d)
        except ValueError:
            return len(CANONICAL_DEGREE_ORDER)

    return sorted(seen.keys(), key=sort_key)


# ==================================================
# DATE HELPERS  — always output YYYY-MM-DD
# ==================================================
DATE_FORMATS = [
    "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%y",
    "%d-%B-%Y", "%d %B, %Y", "%d %b, %Y",
    "%B %d %Y", "%b %d %Y",
    "%d-%m-%Y", "%Y/%m/%d",
]


def parse_deadline(s: str) -> date | None:
    s = re.sub(r"\s+", " ", s).strip()
    if not s:
        return None

    # Strip ordinal suffixes: 1st, 2nd, 3rd, 4th … 31st
    s_clean = re.sub(r"(\d+)(st|nd|rd|th)\b", r"\1", s, flags=re.IGNORECASE)

    for src in (s_clean, s):
        for fmt in DATE_FORMATS:
            try:
                return datetime.strptime(src, fmt).date()
            except ValueError:
                pass

    # Regex fallback: "15 March 2026" anywhere in the string
    m = re.search(
        r"(\d{1,2})\s+"
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})",
        s, re.IGNORECASE,
    )
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)} {m.group(2).title()} {m.group(3)}", "%d %B %Y"
            ).date()
        except ValueError:
            pass

    # Regex fallback: "March 15, 2026"
    m = re.search(
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
        s, re.IGNORECASE,
    )
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


def format_deadline(s: str) -> str:
    """Return ISO YYYY-MM-DD or the original string if unparseable."""
    d = parse_deadline(s)
    return d.isoformat() if d else s.strip()


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
            print(f"  📝  Header row written to '{SHEET_TAB}'")
            return set()
        _header_written = True
        headers = rows[0]
        try:
            col_idx = headers.index("Scholarship Link")
        except ValueError:
            return set()
        return {
            rows[i][col_idx].strip()
            for i in range(1, len(rows))
            if len(rows[i]) > col_idx and rows[i][col_idx].strip()
        }
    except Exception as e:
        print(f"  ⚠  Could not load existing links from Sheets: {e}")
        return set()


def flush_to_sheets(records: list[dict], retries: int = 3):
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
                [str(rec.get(col, "") or "") for col in COLUMN_HEADERS]
                for rec in records
            ]
            ws.append_rows(rows, value_input_option="USER_ENTERED")
            return

        except gspread.exceptions.APIError as e:
            if "RESOURCE_EXHAUSTED" in str(e) or "429" in str(e):
                wait = (attempt + 1) * 15
                print(f"  ⏳  Sheets rate limit — waiting {wait}s …")
                time.sleep(wait)
            else:
                print(f"  ⚠  Sheets API error: {e}")
                return
        except Exception as e:
            print(f"  ⚠  flush_to_sheets error: {e}")
            time.sleep(5)


# ==================================================
# JS EXTRACTORS
# ==================================================

# --- Listing page: grab ALL article cards ---
# Tries multiple known Elementor card selectors; falls back to any <article>
CARDS_JS = r"""
() => {
    const results = [];
    const seen    = new Set();

    // Primary: Elementor grid posts
    document.querySelectorAll(
        'article.elementor-post, article.elementor-grid-item, .elementor-posts article'
    ).forEach(card => {
        let title = '';
        const titleEl = card.querySelector(
            '.elementor-post__title a, .elementor-post__title, h2 a, h3 a'
        );
        if (titleEl) title = (titleEl.innerText || '').trim();

        let link = '';
        // Prefer explicit "read more" links
        const readMore = card.querySelector(
            'a.elementor-post__read-more, a[class*="read-more"], a[class*="readmore"]'
        );
        if (readMore) link = readMore.getAttribute('href') || '';
        if (!link) {
            const thumb = card.querySelector('a.elementor-post__thumbnail__link');
            if (thumb) link = thumb.getAttribute('href') || '';
        }
        if (!link) {
            const titleA = card.querySelector('.elementor-post__title a, h2 a, h3 a');
            if (titleA) link = titleA.getAttribute('href') || '';
        }

        if (title && link && !seen.has(link)) {
            seen.add(link);
            results.push({ title, link });
        }
    });

    // Fallback: any <article> not already caught
    if (results.length === 0) {
        document.querySelectorAll('article').forEach(card => {
            const titleA = card.querySelector('h2 a, h3 a, h1 a');
            if (!titleA) return;
            const title = (titleA.innerText || '').trim();
            const link  = titleA.getAttribute('href') || '';
            if (title && link && !seen.has(link)) {
                seen.add(link);
                results.push({ title, link });
            }
        });
    }

    return results;
}
"""

CONTENT_JS = r"""
() => {
    const junk = 'nav, header, footer, form, .breadcrumb, .pagination, ' +
                 '.share, .social, .ezoic-ad, [id*="ezoic"], script, style';
    const strip = el => {
        el.querySelectorAll(junk).forEach(n => n.remove());
        return (el.innerText || '').trim();
    };
    const selectors = [
        '.elementor-widget-theme-post-content',
        '.entry-content',
        'article .post-content',
        'main article',
        'article',
        'main',
    ];
    for (const sel of selectors) {
        const el = document.querySelector(sel);
        if (el) {
            const clone = el.cloneNode(true);
            const text  = strip(clone);
            if (text.length > 200) return text;
        }
    }
    let best = '';
    document.querySelectorAll('div, section').forEach(el => {
        const t = (el.innerText || '').trim();
        if (t.length > best.length) best = t;
    });
    return best;
}
"""

# ==================================================
# GENERIC HELPERS
# ==================================================
def safe_str(val) -> str:
    if val is None:
        return ""
    if isinstance(val, (list, tuple)):
        return ", ".join(str(v).strip() for v in val if v)
    return str(val).strip()


def clean(s: str) -> str:
    s = safe_str(s)
    for old, new in {
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
        "\u2013": "-", "\u2014": "-", "\u2026": "...", "\u00a0": " ",
    }.items():
        s = s.replace(old, new)
    return re.sub(r"\s+", " ", s).strip()


# ==================================================
# ROUTE BLOCKER
# ==================================================
async def block_junk(route):
    if route.request.resource_type in BLOCKED_TYPES:
        await route.abort()
    else:
        await route.continue_()


# ==================================================
# STAGE 1 — DISCOVER ALL LINKS
# Retries each page up to 3 times; logs card count per page.
# ==================================================
async def discover_all(browser) -> list[dict]:
    scholarships: list[dict] = []
    seen: set[str] = set()

    ctx  = await browser.new_context(extra_http_headers=HEADERS)
    await ctx.route("**/*", block_junk)
    page = await ctx.new_page()
    page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

    for page_num in range(1, TOTAL_PAGES + 1):
        url = BASE_URL if page_num == 1 else PAGE_URL_TMPL.format(n=page_num)
        print(f"\n  📄  Page {page_num}/{TOTAL_PAGES}  →  {url}")

        cards = []
        for attempt in range(3):
            try:
                await page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
                # Wait for Elementor posts OR any article (broader fallback)
                try:
                    await page.wait_for_selector(
                        "article.elementor-post, article.elementor-grid-item, article",
                        timeout=25_000,
                    )
                except Exception:
                    pass
                # Give JS frameworks a moment to render
                await page.wait_for_timeout(1_200)
                cards = await page.evaluate(CARDS_JS)
                if cards:
                    break
                print(f"      ↩  attempt {attempt+1}: 0 cards, retrying …")
                await asyncio.sleep(2)
            except Exception as e:
                print(f"  ⚠   Page {page_num} attempt {attempt+1} failed: {e}")
                await asyncio.sleep(3)

        if not cards:
            print(f"  ⚠   Page {page_num} — no cards found after 3 attempts, skipping")
            continue

        new_count = 0
        for item in cards:
            title = clean(item["title"])
            link  = item["link"].strip()
            if not link:
                continue
            if not link.startswith("http"):
                link = "https://scholarshipscorner.website" + link
            if link in seen:
                continue
            seen.add(link)
            scholarships.append({"name": title, "link": link})
            new_count += 1

        print(f"      +{new_count} cards  (page total: {len(cards)}, "
              f"running total: {len(scholarships)})")

    await ctx.close()
    return scholarships


# ==================================================
# STAGE 2a — EXTRACT OFFICIAL WEBSITE LINK
# ==================================================
async def get_official_link(page) -> str:
    for sel in [
        "a.buttons.btn_green",
        "a.buttons.center",
        "div.button-center a.buttons",
        "a.btn_green",
        ".wp-block-buttons a",
    ]:
        try:
            loc = page.locator(sel).first
            if await loc.count() > 0:
                href = await loc.get_attribute("href")
                if href and href.startswith("http"):
                    return href
        except Exception:
            pass

    for text in ["Visit Official Website", "Official Website",
                 "Apply Now", "Apply Here", "Apply"]:
        try:
            loc = page.locator(f"a:has-text('{text}')").first
            if await loc.count() > 0:
                href = await loc.get_attribute("href")
                if href and href.startswith("http") and "scholarshipscorner" not in href:
                    return href
        except Exception:
            pass

    return ""


# ==================================================
# STAGE 2b — AI FIELD EXTRACTION
# ==================================================
def build_prompt(text: str, url: str, title: str) -> str:
    return (
        "You are a precise data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON — no markdown fences, no commentary.\n"
        "- All values must be plain ASCII strings (no arrays, no null).\n"
        "- Use straight quotes only.\n"
        "- degree_type: comma-separated list. Allowed values ONLY: "
        "Bachelors, Masters, MBA, Executive MBA, PhD, Postdoctoral.\n"
        "  If multiple degrees are offered list them all: e.g. 'Bachelors, Masters, PhD'.\n"
        "- scholarship_type: e.g. Fully Funded, Partial, Government, University.\n"
        "- deadline: exact date string as found in the text (e.g. '31 March 2026').\n"
        "  If multiple deadlines exist, use the earliest upcoming one.\n"
        "- host_country: country where the student will physically study.\n"
        "- field_of_study: comma-separated fields if multiple.\n\n"
        "JSON schema:\n"
        "{\n"
        '  "scholarship_type": "",\n'
        '  "deadline": "",\n'
        '  "host_country": "",\n'
        '  "degree_type": "",\n'
        '  "field_of_study": ""\n'
        "}\n\n"
        f"Page URL: {url}\n"
        f"Page Title: {title}\n\n"
        f"Text:\n{text[:8000]}"
    )


def parse_json(raw: str) -> dict:
    raw = re.sub(r"^```[a-z]*\s*", "", raw.strip(), flags=re.I)
    raw = re.sub(r"\s*```$", "", raw).strip()
    try:
        d = json.loads(raw)
        return d if isinstance(d, dict) else dict(EMPTY_SCHEMA)
    except Exception:
        m = re.search(r"\{.*\}", raw, flags=re.DOTALL)
        if m:
            try:
                d = json.loads(m.group(0))
                return d if isinstance(d, dict) else dict(EMPTY_SCHEMA)
            except Exception:
                pass
    return dict(EMPTY_SCHEMA)


async def ai_extract(text: str, url: str, title: str) -> dict:
    try:
        resp = await ai_client.chat.completions.create(
            model="gpt-4o-mini",
            temperature=0,
            messages=[
                {"role": "system",
                 "content": "You are a structured data extraction engine. Return only valid JSON."},
                {"role": "user", "content": build_prompt(text, url, title)},
            ],
            max_tokens=500,
        )
        raw = resp.choices[0].message.content or ""
    except Exception as e:
        print(f"  ⚠  OpenAI error: {e}")
        return dict(EMPTY_SCHEMA)

    data = parse_json(raw)
    out  = dict(EMPTY_SCHEMA)
    for k in out:
        if k in data and data[k] is not None:
            out[k] = clean(str(data[k]))
    return out


# ==================================================
# STAGE 2 — SCRAPE ONE DETAIL PAGE
# Returns a LIST of rows (one per degree type).
# ==================================================
async def scrape_one(
    s:          dict,
    browser,
    scrape_sem: asyncio.Semaphore,
    ai_sem:     asyncio.Semaphore,
    idx:        int,
    total:      int,
) -> list[dict] | None:
    """
    Returns:
        list[dict]  — one or more sheet rows (one per degree type)
        None        — skip this scholarship (expired / load error)
    """
    name, link = s["name"], s["link"]

    async with scrape_sem:
        ctx  = await browser.new_context(extra_http_headers=HEADERS)
        await ctx.route("**/*", block_junk)
        page = await ctx.new_page()
        page.set_default_navigation_timeout(NAV_TIMEOUT_MS)

        try:
            await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT_MS)
            await page.wait_for_timeout(800)
        except Exception as e:
            print(f"  ⚠  ({idx}/{total}) Load failed — {name[:50]}: {e}")
            await ctx.close()
            return None

        text       = clean(await page.evaluate(CONTENT_JS))
        official   = await get_official_link(page)
        page_title = clean(await page.title())
        await ctx.close()

    async with ai_sem:
        data = await ai_extract(text, link, page_title)

    # ── Deadline check ──────────────────────────────────────────────────
    deadline_raw = clean(data.get("deadline", ""))

    if deadline_raw and deadline_passed(deadline_raw):
        print(f"  ⏭  ({idx}/{total}) SKIP (expired: {deadline_raw})  {name[:50]}")
        return None

    deadline_out = format_deadline(deadline_raw)   # always YYYY-MM-DD or raw

    # ── Degree expansion ────────────────────────────────────────────────
    degree_raw  = clean(data.get("degree_type", ""))
    degree_list = normalise_degrees(degree_raw)

    # If no degree extracted, use a single row with blank degree_type
    if not degree_list:
        degree_list = [""]

    base_row = {
        "Scholarship Name": clean(name),
        "Scholarship Link": link,
        "Official Website": official,
        "Scholarship Type": clean(data.get("scholarship_type", "")),
        "Deadline":         deadline_out,
        "Host Country":     clean(data.get("host_country", "")),
        "Field of Study":   clean(data.get("field_of_study", "")),
    }

    rows = []
    for deg in degree_list:
        row = dict(base_row)
        row["Degree Type"] = deg
        rows.append(row)

    label = " | ".join(d for d in degree_list if d) or "(no degree)"
    print(f"  ✅  ({idx}/{total}) [{len(rows)} row(s) → {label}]  {name[:55]}")
    return rows


# ==================================================
# MAIN
# ==================================================
async def main():
    print("=" * 66)
    print("  ScholarshipsCorner Scraper  →  Google Sheets  [FINAL v2]")
    print(f"  Pages  : {TOTAL_PAGES}")
    print(f"  Sheet  : {SHEET_ID}")
    print(f"  Tab    : {SHEET_TAB}")
    print(f"  Today  : {TODAY.isoformat()}")
    print("=" * 66)

    print("\n  🔗  Connecting to Google Sheets …")
    existing = load_existing_links()
    print(f"  🗃   {len(existing)} already-scraped links loaded\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)

        print("🌐  Stage 1 — Discovering all scholarship links …\n")
        all_scholarships = await discover_all(browser)
        print(f"\n  📦  {len(all_scholarships)} unique scholarships found")

        todo = [s for s in all_scholarships if s["link"] not in existing]
        print(f"  ⏩  {len(all_scholarships) - len(todo)} already in Sheets — skipping")
        print(f"  🆕  {len(todo)} new scholarships to scrape\n")

        if not todo:
            print("  Nothing new. Exiting.")
            await browser.close()
            return

        print("🧭  Stage 2 — Scraping detail pages …\n")
        scrape_sem = asyncio.Semaphore(MAX_CONCURRENT_BROWSERS)
        ai_sem     = asyncio.Semaphore(MAX_CONCURRENT_OPENAI)

        total        = len(todo)
        saved_links  = 0   # unique scholarship URLs saved
        saved_rows   = 0   # total sheet rows written (>= saved_links due to expansion)
        skipped      = 0
        buf: list[dict] = []

        for batch_start in range(0, total, FLUSH_EVERY):
            batch = todo[batch_start : batch_start + FLUSH_EVERY]

            tasks = [
                scrape_one(s, browser, scrape_sem, ai_sem,
                           batch_start + j + 1, total)
                for j, s in enumerate(batch)
            ]
            results = await asyncio.gather(*tasks, return_exceptions=True)

            for r in results:
                if isinstance(r, list):
                    buf.extend(r)
                    # Mark the base URL as seen so duplicates are skipped next run
                    existing.add(r[0]["Scholarship Link"])
                    saved_links += 1
                    saved_rows  += len(r)
                elif r is None:
                    skipped += 1
                elif isinstance(r, Exception):
                    print(f"  ⚠  Task error: {r}")
                    skipped += 1

            if buf:
                flush_to_sheets(buf)
                print(
                    f"\n  💾  Flushed {len(buf)} rows to Sheets  "
                    f"[scholarships={saved_links} | rows={saved_rows} | skipped={skipped}]\n"
                )
                buf.clear()

        await browser.close()

    print("\n" + "=" * 66)
    print("  ✅  DONE")
    print(f"  Scholarships saved : {saved_links}")
    print(f"  Sheet rows written : {saved_rows}  (expanded by degree type)")
    print(f"  Skipped            : {skipped}  (expired deadline / load error)")
    print(f"  Sheet  : https://docs.google.com/spreadsheets/d/{SHEET_ID}")
    print("=" * 66)


if __name__ == "__main__":
    asyncio.run(main())