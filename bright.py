"""
BrightScholarship.com Scraper  →  Google Sheets Version
=========================================================
Site   : https://brightscholarship.com/category/scholarships/
Output : Google Sheets → Tab: bright_scholarships

NOTE: This scraper has different columns than others:
      "Offered By" + "Eligible Nations" columns
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
BASE_URL      = "https://brightscholarship.com/category/scholarships/"
PAGE_URL_TMPL = "https://brightscholarship.com/category/scholarships/page/{n}/"
START_PAGE    = 1
MAX_PAGES     = 81
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
SHEET_TAB        = "bright_scholarships"

# BrightScholarship unique columns: "Offered By" + "Eligible Nations"
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
# DEGREE NORMALISATION
# ══════════════════════════════════════════════════════════════
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
    # PostDoc
    "postdoctoral":     "PostDoc",
    "postdoc":          "PostDoc",
    "post-doctoral":    "PostDoc",
    "post doc":         "PostDoc",
}

CANONICAL_DEGREE_ORDER = ["Bachelors", "Masters", "MBA", "Executive MBA", "PhD", "PostDoc"]


def normalise_degrees(raw: str) -> list[str]:
    """
    Split a raw degree string and return a deduplicated list of
    canonical degree labels (e.g. ['Bachelors', 'Masters', 'PhD']).
    Unknown tokens are kept as-is (title-cased).
    Returns [''] when nothing is found so callers always get one row.
    """
    if not raw:
        return [""]

    tokens = re.split(r"[,;/|&]+|\band\b|\bor\b|\bwith\b", raw, flags=re.IGNORECASE)
    seen: dict[str, bool] = {}

    for tok in tokens:
        tok = tok.strip().rstrip(".")
        if not tok:
            continue
        key = tok.lower()

        if key in DEGREE_NORM:
            canon = DEGREE_NORM[key]
        else:
            canon = None
            for pattern, label in DEGREE_NORM.items():
                if pattern in key:
                    canon = label
                    break
            if canon is None:
                canon = tok.title()

        seen[canon] = True

    def sort_key(d):
        try:
            return CANONICAL_DEGREE_ORDER.index(d)
        except ValueError:
            return len(CANONICAL_DEGREE_ORDER)

    result = sorted(seen.keys(), key=sort_key)
    return result if result else [""]


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


def save_rows_to_sheets(records: list[dict], retries: int = 3):
    """Save one or more rows for a single scholarship (expanded by degree type)."""
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
                print(f"  ⏳  Rate limit — waiting {wait}s ...")
                time.sleep(wait)
            else:
                print(f"  ⚠  Sheets API error: {e}")
                return
        except Exception as e:
            print(f"  ⚠  save_rows error: {e}")
            time.sleep(5)


# ══════════════════════════════════════════════════════════════
# JAVASCRIPT
# ══════════════════════════════════════════════════════════════
EXTRACT_CARDS_JS = r"""
() => {
    const results = [];
    const seen    = new Set();

    document.querySelectorAll('a.td-image-wrap[href][title]').forEach(a => {
        const href  = (a.getAttribute('href')  || '').trim();
        const title = (a.getAttribute('title') || '').trim();
        if (!href || !title || seen.has(href)) return;
        if (!href.includes('brightscholarship.com')) return;

        const low = href.toLowerCase();
        const skipPatterns = [
            'how-to-', 'what-is-', 'complete-guide', 'cover-letter',
            '/resume', '/cv-', 'tips-for', 'best-way', 'internship/',
            'conference/', 'result/', 'online-course', 'jobs/',
        ];
        for (const p of skipPatterns) { if (low.includes(p)) return; }

        const lTitle = title.toLowerCase();
        const isScholarship = (
            low.includes('scholarship') || low.includes('fellowship') ||
            low.includes('funded')      || low.includes('grant')      ||
            low.includes('award')       || lTitle.includes('scholarship') ||
            lTitle.includes('fellowship') || lTitle.includes('funded')
        );
        if (!isScholarship) return;

        seen.add(href);
        results.push({ title, link: href });
    });
    return results;
}
"""

EXTRACT_DETAIL_JS = r"""
() => {
    const fields = {};

    document.querySelectorAll('table tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td'));
        if (cells.length >= 2) {
            const label = (cells[0].innerText || '').trim();
            const value = (cells[1].innerText || '').trim();
            if (label && value) fields[label.toLowerCase()] = value;
        }
    });

    document.querySelectorAll('.wp-block-table tr, .kb-table tr').forEach(row => {
        const cells = Array.from(row.querySelectorAll('td, th'));
        if (cells.length >= 2) {
            const label = (cells[0].innerText || '').trim();
            const value = (cells[1].innerText || '').trim();
            if (label && value) fields[label.toLowerCase()] = value;
        }
    });

    let officialUrl = '';
    document.querySelectorAll('a[target="_blank"]').forEach(a => {
        if (officialUrl) return;
        const href    = (a.getAttribute('href') || '').trim();
        const btnSpan = a.querySelector('span.kt-btn-inner-text, span[class*="inner-text"]');
        const btnText = btnSpan
            ? (btnSpan.innerText || '').trim().toLowerCase()
            : (a.innerText || '').trim().toLowerCase();
        if (
            href.startsWith('http') &&
            !href.includes('brightscholarship.com') &&
            (btnText === 'official link' || btnText === 'apply link' ||
             btnText === 'apply now'     || btnText.includes('official link') ||
             btnText.includes('apply link'))
        ) { officialUrl = href; }
    });

    const article = document.querySelector(
        '.td-post-content, .entry-content, article .tdc-content-wrap, article'
    );
    const clone = article ? article.cloneNode(true) : document.body.cloneNode(true);
    clone.querySelectorAll(
        'script, style, nav, header, footer, .sharedaddy, ' +
        '.td-related-title, .td-post-next-prev, .td-author-box, ' +
        '[class*="comment"], [id*="ads"], [class*="adsbygoogle"]'
    ).forEach(n => n.remove());
    const text = (clone.innerText || '').trim();

    return { fields, officialUrl, text };
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
    if not s or s.lower() in ("none", "n/a", "-", "open", "ongoing", ""):
        return None
    formats = [
        "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y",
        "%d-%B-%Y", "%d-%b-%Y", "%d %B, %Y", "%d %b, %Y",
        "%B %Y", "%b %Y",
    ]
    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)[,\s]+(\d{4})", s, re.I)
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


def get_field(fields: dict, *keys) -> str:
    for key in keys:
        for fk, fv in fields.items():
            if key.lower() in fk.lower():
                val = clean(fv)
                if val:
                    return val
    return ""


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(name: str, fields: dict, text: str) -> str:
    table_rows = "\n".join(f"  {k}: {v}" for k, v in fields.items()) if fields else "  (none found)"
    return (
        "You are a precise scholarship data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON — no markdown, no commentary.\n"
        "- All values must be plain ASCII strings (never arrays or null).\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated from: Bachelors, Masters, MBA, PhD, PostDoc.\n"
        "- scholarship_type: e.g. 'Fully Funded', 'Partial Funding', 'Tuition Fee waiver'.\n"
        "- deadline: exact date string as found (e.g. '10 April 2026'). '' if not found.\n"
        "- host_country: the country where the student will STUDY.\n"
        "- field_of_study: comma-separated subject fields if mentioned.\n\n"
        '{"scholarship_type":"","deadline":"","host_country":"","degree_type":"","field_of_study":""}\n\n'
        f"Scholarship Name: {name}\n\nDetail table:\n{table_rows}\n\nFull page text:\n{text[:7000]}"
    )


async def ai_extract(name: str, fields: dict, text: str) -> dict:
    try:
        resp = await ai_client.chat.completions.create(
            model="gpt-4o-mini", temperature=0,
            messages=[
                {"role": "system",
                 "content": "Scholarship data extraction engine. Return only valid JSON."},
                {"role": "user", "content": build_prompt(name, fields, text)},
            ],
            max_tokens=400,
        )
        raw = resp.choices[0].message.content or ""
    except Exception as e:
        print(f"  ⚠  OpenAI error: {e}")
        return dict(EMPTY)

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


# ══════════════════════════════════════════════════════════════
# SCRAPE ONE CARD
# Returns a list of rows (one per degree type) or None to skip.
# ══════════════════════════════════════════════════════════════

async def scrape_one(page, card: dict, seq: int, total: int) -> list[dict] | None:
    name = clean(card.get("title", ""))
    link = card["link"]

    print(f"\n  ── ({seq}/{total}) {name[:58]}")
    print(f"     link : {link}")

    try:
        await page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await page.wait_for_timeout(DETAIL_WAIT)
    except Exception as e:
        print(f"  ⚠  Load failed: {e}")
        return None

    data         = await page.evaluate(EXTRACT_DETAIL_JS)
    fields       = data.get("fields",      {})
    official_url = clean(data.get("officialUrl", ""))
    text         = clean(data.get("text",  ""))

    print(f"     fields={len(fields)}  official={'✓' if official_url else '✗'}  text={len(text)}chars")

    offered_by   = get_field(fields, "offered by",    "provider",    "organization", "host")
    degree_raw   = get_field(fields, "degree level",  "degree",      "eligible degree", "level")
    coverage     = get_field(fields, "scholarship coverage", "coverage", "funding type", "scholarship type", "award")
    nationality  = get_field(fields, "eligible nationality", "nationality", "eligible for", "open to")
    host_country = get_field(fields, "award country", "host country", "country", "location", "taken at", "study in")
    deadline_raw = get_field(fields, "last date",     "deadline",    "closing date", "apply before", "due date", "last day")

    if not offered_by:
        for pat in [
            r'((?:University|Institut(?:e|o|ut)|College|School|Academy|Hochschule|Universit[äae]t)\s+(?:of\s+)?[A-Z][a-zA-Z\s\-]{2,40}?)(?:\s+(?:Scholarship|Fellowship|Award|Grant|2\d{3}|in\b))',
            r'([A-Z][a-zA-Z\s\-]{2,40}?\s+(?:University|Institute|College|School|Academy))(?:\s+(?:Scholarship|Fellowship|Award|Grant|2\d{3}|in\b))',
        ]:
            m = re.search(pat, name)
            if m:
                offered_by = m.group(1).strip()
                break

    print(f"     offered={offered_by or '—'}  country={host_country or '—'}  degree={degree_raw or '—'}  dl={deadline_raw or '—'}")

    if deadline_raw and deadline_passed(deadline_raw):
        print(f"  ⏭  SKIP — deadline passed [{deadline_raw}]")
        return None

    deadline_out = ""
    if deadline_raw:
        parsed = parse_deadline(deadline_raw)
        deadline_out = parsed.isoformat() if parsed else deadline_raw

    ai_data = await ai_extract(name, fields, text)

    final_degree_raw = clean_degree(degree_raw or clean(str(ai_data.get("degree_type", ""))))
    final_type       = coverage              or clean(str(ai_data.get("scholarship_type", "")))
    final_country    = host_country          or clean(str(ai_data.get("host_country",     "")))
    final_field      = clean(str(ai_data.get("field_of_study", "")))

    if not deadline_out:
        ai_dl = clean(str(ai_data.get("deadline", "")))
        if ai_dl:
            parsed = parse_deadline(ai_dl)
            if parsed and not deadline_passed(ai_dl):
                deadline_out = parsed.isoformat()

    print(f"  ✅  dl=[{deadline_out or 'blank'}]  web=[{'✓' if official_url else '✗'}]  country=[{final_country or '—'}]")

    # ── Degree expansion: one row per degree type ──────────────────────
    degree_list = normalise_degrees(final_degree_raw)

    base_row = {
        "Scholarship Name": name,
        "Scholarship Link": link,
        "Offered By":       offered_by,
        "Official Website": official_url,
        "Scholarship Type": final_type,
        "Host Country":     final_country,
        "Eligible Nations": nationality,
        "Deadline":         deadline_out,
        "Field of Study":   final_field,
    }

    rows = []
    for deg in degree_list:
        row = dict(base_row)
        row["Degree Type"] = deg
        rows.append(row)

    label = " | ".join(d for d in degree_list if d) or "(no degree)"
    print(f"     → {len(rows)} row(s) expanded: {label}")

    return rows


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═" * 66)
    print("  BrightScholarship.com Scraper  →  Google Sheets")
    print(f"  Pages  : {START_PAGE} → {MAX_PAGES}")
    print(f"  Tab    : {SHEET_TAB}")
    print("═" * 66)

    print("\n  🔗  Connecting to Google Sheets ...")
    existing = load_existing_links()
    print(f"\n  🗃  {len(existing)} already-saved links loaded\n")

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
            if route.request.resource_type in {"image", "font", "media"}:
                await route.abort()
            else:
                await route.continue_()

        listing_page = await ctx.new_page()
        detail_page  = await ctx.new_page()
        await listing_page.route("**/*", block_junk)
        await detail_page.route("**/*",  block_junk)
        listing_page.set_default_navigation_timeout(NAV_TIMEOUT)
        detail_page.set_default_navigation_timeout(NAV_TIMEOUT)

        saved_scholarships = 0   # unique scholarship URLs saved
        saved_rows         = 0   # total sheet rows (>= scholarships due to expansion)
        skipped            = 0

        for page_num in range(START_PAGE, MAX_PAGES + 1):
            url = BASE_URL if page_num == 1 else PAGE_URL_TMPL.format(n=page_num)
            print(f"\n{'═'*66}")
            print(f"  📄  PAGE {page_num}/{MAX_PAGES}  →  {url}")
            print(f"{'═'*66}")

            try:
                await listing_page.goto(url, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
                await listing_page.wait_for_timeout(PAGE_WAIT)
            except Exception as e:
                print(f"  ⚠  Page {page_num} failed: {e} — skipping")
                continue

            cards = await listing_page.evaluate(EXTRACT_CARDS_JS)
            if not cards:
                print(f"  ⚠  No cards on page {page_num} — stopping.")
                break

            new_cards = [c for c in cards if c["link"] not in existing]
            print(f"  Found {len(cards)}  |  {len(cards)-len(new_cards)} saved  |  {len(new_cards)} to scrape\n")

            for i, card in enumerate(new_cards, start=1):
                existing.add(card["link"])
                rows = await scrape_one(detail_page, card, seq=i, total=len(new_cards))

                if rows:
                    save_rows_to_sheets(rows)
                    saved_scholarships += 1
                    saved_rows         += len(rows)
                    print(f"  💾  Saved {len(rows)} row(s) to Sheets  "
                          f"[scholarships={saved_scholarships} | rows={saved_rows}]")
                else:
                    skipped += 1

                await asyncio.sleep(0.5)

        await browser.close()

    print("\n" + "═" * 66)
    print("  ✅  DONE")
    print(f"  Scholarships saved : {saved_scholarships}")
    print(f"  Sheet rows written : {saved_rows}  (expanded by degree type)")
    print(f"  Skipped            : {skipped}  (expired / load error)")
    print(f"  Sheet  : https://docs.google.com/spreadsheets/d/{SHEET_ID}")
    print("═" * 66)


if __name__ == "__main__":
    asyncio.run(main())