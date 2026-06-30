"""
WeMakeScholars Scholarship Scraper  →  Google Sheets Version
==============================================================
Site   : https://www.wemakescholars.com/scholarship?nationality=83
Output : Google Sheets → Tab: wms_scholarships

NOTE: This scraper has an extra "Provider Name" column
      compared to other scrapers — kept as-is.
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
START_URL         = "https://www.wemakescholars.com/scholarship?nationality=83"
NAV_TIMEOUT       = 60_000
PAGE_WAIT         = 3000
DETAIL_WAIT       = 2500
LOAD_MORE_WAIT    = 3000
MAX_AI_CONCURRENT = 6
TODAY             = date.today()

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
SHEET_TAB        = "wms_scholarships"

# WeMakeScholars has "Provider Name" — unique to this scraper
COLUMN_HEADERS = [
    "Scholarship Name", "Scholarship Link", "Provider Name",
    "Official Website", "Scholarship Type", "Deadline",
    "Host Country", "Degree Type", "Field of Study",
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
# JAVASCRIPT
# ══════════════════════════════════════════════════════════════
COLLECT_CARDS_JS = r"""
() => {
    const results = [];
    const seen    = new Set();
    document.querySelectorAll('#scholarship-content .post').forEach(post => {
        let name = '', link = '';
        const titleA = post.querySelector('h2 a, h3 a, .clrwms a, a.clrwms, a[href*="/scholarship/"]');
        if (titleA) {
            name = (titleA.innerText || '').trim();
            link = titleA.getAttribute('href') || '';
            if (link && !link.startsWith('http'))
                link = 'https://www.wemakescholars.com' + link;
        }
        if (!name || seen.has(link)) return;
        seen.add(link);

        let deadline = '', degree = '', funding = '', courses = '', nationality = '', location = '';
        post.querySelectorAll('.text-line-div').forEach(row => {
            const label = (row.querySelector('p.text-line, .text-line') || {}).innerText || '';
            const value = (row.querySelector('span.text-line-value, .text-line-value') || {}).innerText || '';
            const lLow  = label.trim().toLowerCase();
            const vTrim = value.trim();
            if (!vTrim) return;
            if (lLow.includes('deadline'))             deadline    = vTrim;
            else if (lLow.includes('eligible degree')) degree      = vTrim;
            else if (lLow.includes('funding type'))    funding     = vTrim;
            else if (lLow.includes('eligible course')) courses     = vTrim;
            else if (lLow.includes('nationalit'))      nationality = vTrim;
            else if (lLow.includes('taken at') || lLow.includes('location')) location = vTrim;
        });
        results.push({ name, link, deadline, degree, funding, courses, nationality, location });
    });
    return results;
}
"""

EXTRACT_DETAIL_JS = r"""
() => {
    const fields = {};
    document.querySelectorAll('.text-line-div').forEach(row => {
        const label = ((row.querySelector('p.text-line') || {}).innerText || '').trim().toLowerCase();
        const val   = ((row.querySelector('span.text-line-value') || {}).innerText || '').trim();
        if (label && val) fields[label] = val;
    });

    let uniPageLink = '', uniName = '';
    document.querySelectorAll('.text-line-div').forEach(row => {
        const label = ((row.querySelector('p.text-line') || {}).innerText || '').trim().toLowerCase();
        if (label.includes('provider')) {
            const a = row.querySelector('span.text-line-value a.clrblue, span.text-line-value a[href*="/university/"]');
            if (a) {
                uniName = (a.innerText || '').trim();
                const href = (a.getAttribute('href') || '').trim();
                uniPageLink = href.startsWith('http') ? href : 'https://www.wemakescholars.com' + href;
            }
        }
    });

    let officialUrl = '';
    const disclaimer = document.querySelector('#disclaimer');
    if (disclaimer) {
        const sourceLink = disclaimer.querySelector('a[target="_blank"][rel*="nofollow"], a[target="_blank"]');
        if (sourceLink) {
            const href = (sourceLink.getAttribute('href') || '').trim();
            if (href.startsWith('http') && !href.includes('wemakescholars')) officialUrl = href;
        }
    }

    const main = document.querySelector('.sub-post, .panel.clearfix, main, #main-layout-content');
    const clone = main ? main.cloneNode(true) : document.body.cloneNode(true);
    clone.querySelectorAll('script,style,nav,header,footer,.comments-panel,[class*="widget"],[id*="banner"]')
         .forEach(n => n.remove());
    const text = (clone.innerText || '').trim();

    return { fields, uniPageLink, uniName, officialUrl, text };
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
        "%d %b, %Y", "%d %B, %Y", "%d %b %Y", "%d %B %Y",
        "%B %d, %Y", "%b %d, %Y", "%Y-%m-%d", "%d/%m/%Y",
        "%m/%d/%Y",  "%d-%b-%Y",  "%d-%B-%Y", "%B %Y", "%b %Y",
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
    return None


def deadline_passed(s: str) -> bool:
    d = parse_deadline(s)
    return d is not None and d < TODAY


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(text: str, name: str, card: dict, fields: dict) -> str:
    known = "\n".join(f"  {k}: {v}" for k, v in fields.items()) if fields else "N/A"
    return (
        "You are a precise data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON — no markdown, no commentary.\n"
        "- All values must be plain ASCII strings.\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated from: Bachelors, Masters, MBA, PhD, Executive MBA.\n"
        "- scholarship_type: e.g. 'Fully Funded', 'Partial Funding', 'Fee waiver'.\n"
        "- deadline: exact date as found. Return '' if not mentioned.\n"
        "- host_country: country where student will study.\n"
        "- field_of_study: comma-separated subject fields.\n\n"
        '{"scholarship_type":"","deadline":"","host_country":"","degree_type":"","field_of_study":""}\n\n'
        f"Name: {name}\nCard:\n  deadline:{card.get('deadline','')}\n  degree:{card.get('degree','')}\n"
        f"  funding:{card.get('funding','')}\n  location:{card.get('location','')}\n\n"
        f"Detail fields:\n{known}\n\nText:\n{text[:6000]}"
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
                     fields: dict, sem: asyncio.Semaphore) -> dict:
    async with sem:
        try:
            resp = await ai_client.chat.completions.create(
                model="gpt-4o-mini", temperature=0,
                messages=[
                    {"role": "system", "content": "Structured data extraction. Return only valid JSON."},
                    {"role": "user",   "content": build_prompt(text, name, card, fields)},
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
    detail_page, card: dict, seq: int, total: int, ai_sem: asyncio.Semaphore,
) -> dict | None:

    name         = clean(card["name"])
    link         = card["link"]
    deadline_raw = clean(card.get("deadline",  ""))
    degree_card  = clean(card.get("degree",    ""))
    funding_card = clean(card.get("funding",   ""))
    location     = clean(card.get("location",  ""))

    if deadline_raw and deadline_passed(deadline_raw):
        print(f"  ⏭  ({seq}/{total}) SKIP expired [{deadline_raw}] — {name[:45]}")
        return None

    parsed_d     = parse_deadline(deadline_raw)
    deadline_out = parsed_d.isoformat() if parsed_d else ""

    print(f"\n  🔎  ({seq}/{total}) {name[:55]}")
    print(f"       deadline : {deadline_raw or 'not on card'}")

    try:
        await detail_page.goto(link, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await detail_page.wait_for_timeout(DETAIL_WAIT)
    except Exception as e:
        print(f"  ⚠  Detail load failed: {e}")
        return None

    detail       = await detail_page.evaluate(EXTRACT_DETAIL_JS)
    fields       = detail.get("fields",      {})
    uni_page     = detail.get("uniPageLink", "")
    uni_name     = detail.get("uniName",     "")
    official_url = detail.get("officialUrl", "")
    text         = clean(detail.get("text",  ""))

    print(f"       provider : {uni_name or 'NOT FOUND'}")
    print(f"       official : {official_url[:70] or 'NOT FOUND'}")

    if not official_url:
        official_url = uni_page

    ai_data = await ai_extract(text, name, card, fields, ai_sem)

    dl_from_fields = clean(fields.get("deadline", ""))
    final_deadline = deadline_raw or dl_from_fields or clean(ai_data.get("deadline", ""))
    parsed_final   = parse_deadline(final_deadline)
    deadline_out   = parsed_final.isoformat() if parsed_final else ""

    if final_deadline and deadline_passed(final_deadline):
        print(f"  ⏭  SKIP expired (detail) [{final_deadline}]")
        return None

    degree_out   = clean_degree(
        degree_card or clean(fields.get("eligible degrees", "")) or
        clean(ai_data.get("degree_type", ""))
    )
    schol_type   = (
        funding_card or clean(fields.get("funding type", "")) or
        clean(ai_data.get("scholarship_type", ""))
    )
    host_country = (
        location or clean(fields.get("scholarship can be taken at", "")) or
        clean(ai_data.get("host_country", ""))
    )
    field_study  = clean(ai_data.get("field_of_study", ""))

    print(f"  ✅  ({seq}/{total})  {name[:50]}  dl=[{deadline_out or 'blank'}]")

    return {
        "Scholarship Name": name,
        "Scholarship Link": link,
        "Provider Name":    uni_name,
        "Official Website": official_url,
        "Scholarship Type": schol_type,
        "Deadline":         deadline_out,
        "Host Country":     host_country,
        "Degree Type":      degree_out,
        "Field of Study":   field_study,
    }


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═" * 64)
    print("  WeMakeScholars Scraper  →  Google Sheets")
    print(f"  URL    : {START_URL}")
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

        print("🌐  Loading WeMakeScholars ...")
        await listing_page.goto(START_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        await listing_page.wait_for_timeout(PAGE_WAIT)

        ai_sem        = asyncio.Semaphore(MAX_AI_CONCURRENT)
        scraped       = set(existing)
        saved         = 0
        skipped       = 0
        round_n       = 0
        no_new_streak = 0
        buf: list[dict] = []

        print("\n🔄  Starting scrape → load-more → scrape loop ...\n")

        while True:
            round_n  += 1
            all_cards = await listing_page.evaluate(COLLECT_CARDS_JS)
            new_cards = [c for c in all_cards if c["link"] not in scraped]

            if new_cards:
                no_new_streak = 0
                print(f"\n  📋  Round {round_n} — {len(new_cards)} new cards  (seen: {len(scraped)})")

                for i, card in enumerate(new_cards, start=1):
                    scraped.add(card["link"])
                    res = await scrape_one(detail_page, card,
                                           seq=i, total=len(new_cards), ai_sem=ai_sem)
                    if res:
                        buf.append(res); saved += 1
                    else:
                        skipped += 1
                    await listing_page.bring_to_front()

                if buf:
                    flush_to_sheets(buf)
                    print(f"\n  💾  Saved {len(buf)} to Sheets  [total={saved} | skipped={skipped}]\n")
                    buf.clear()

            else:
                no_new_streak += 1
                print(f"  ↕   Round {round_n} — no new cards ({no_new_streak}/3)")
                if no_new_streak >= 3:
                    print("\n  ✅  Done!"); break

            await listing_page.bring_to_front()
            try:
                load_more = listing_page.locator(
                    "input#load-more, input[value*='more Scholarships'], "
                    "input[value*='View more'], #load-more"
                ).first
                if await load_more.count() == 0:
                    print("\n  ✅  No Load More — all loaded!"); break
                if not await load_more.is_visible():
                    print("\n  ✅  Load More hidden — end!"); break
                await load_more.scroll_into_view_if_needed()
                await load_more.click()
                print(f"  🔘  Clicked Load More ...")
                await listing_page.wait_for_timeout(LOAD_MORE_WAIT)
            except Exception as e:
                print(f"\n  ⚠  Load more: {e}")
                print("  ✅  All loaded."); break

        if buf:
            flush_to_sheets(buf)
            print(f"\n  💾  Final flush — {len(buf)} records")

        await browser.close()

    print("\n" + "═" * 64)
    print("  ✅  DONE")
    print(f"  Saved  : {saved}  |  Skipped: {skipped}")
    print(f"  Sheet  : https://docs.google.com/spreadsheets/d/{SHEET_ID}")
    print("═" * 64)


if __name__ == "__main__":
    asyncio.run(main())