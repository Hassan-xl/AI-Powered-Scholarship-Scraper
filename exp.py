"""
Expatrio Study-Finder Scraper  —  Google Sheets Version  [FINAL FIXED]
======================================================================
Site   : https://www.expatrio.com/study-buddy/#/study-finder
Output : Google Sheets → Tab: expatrio_programs

Fixes applied:
  1. DOM precisely targets sibling <span> elements for Deadlines and Tuition using global scope.
  2. AI prompt updated to extract scholarship_type and deadline as a solid fallback.
  3. Hardcoded keys maintained for testing purposes.
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
# ════════════════════════════════════════════════════════���═════
START_URL         = "https://www.expatrio.com/study-buddy/#/study-finder"
SCROLL_STEP       = 800
SCROLL_PAUSE      = 1.5
DETAIL_WAIT       = 5000          # ms to wait after clicking a card
NAV_TIMEOUT       = 90_000
MAX_AI_CONCURRENT = 8
TODAY             = date.today()

# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS CONFIG
# ══════════════════════════════════════════════════════════════
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "expatrio_programs"

COLUMN_HEADERS = [
    "Scholarship Name",   # Program title
    "Scholarship Link",   # N/A (SPA)
    "Official Website",   # University name
    "Scholarship Type",   # No Tuition Fee / Tuition Required
    "Deadline",           # ISO date YYYY-MM-DD or raw text
    "Host Country",
    "Degree Type",
    "Field of Study",
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
    "host_country":     "Germany",
    "degree_type":      "",
    "field_of_study":   "",
}

# ══════════════════════════════════════════════════════════════
# NOISE / JUNK FILTER
# ══════════════════════════════════════════════════════════════
JUNK_TITLES = {
    "close", "back", "login", "sign in", "sign up", "register",
    "continue with google", "continue with apple", "sign in with google",
    "sign in with apple", "add to your list", "save to your list",
    "apply", "details", "none", "all programs", "featured", "my list",
    "find programs", "study field", "degree", "city", "university",
    "duration", "tuition fee", "language",
}

def is_junk_title(title: str) -> bool:
    t = title.strip().lower()
    if not t or len(t) < 3:
        return True
    if t in JUNK_TITLES:
        return True
    if len(t) < 5:
        return True
    return False


# ══════════════════════════════════════════════════════════════
# GOOGLE SHEETS HELPERS
# ══════════════════════════════════════════��═══════════════════
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


def load_existing_names() -> set:
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
            col_idx = headers.index("Scholarship Name")
        except ValueError:
            return set()
        return set(
            rows[i][col_idx].strip()
            for i in range(1, len(rows))
            if len(rows[i]) > col_idx and rows[i][col_idx].strip()
        )
    except Exception as e:
        print(f"  ⚠  Could not load existing names: {e}")
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
# JAVASCRIPT  — Card collector
# ══════════════════════════════════════════════════════════════
COLLECT_CARDS_JS = r"""
() => {
    const halfW = window.innerWidth * 0.65;
    const JUNK = new Set([
        'close','back','login','sign in','sign up','register',
        'continue with google','continue with apple',
        'sign in with google','sign in with apple',
        'add to your list','save to your list','apply','details','none',
        'all programs','featured','my list','find programs',
        'study field','degree','city','university','duration',
        'tuition fee','language'
    ]);

    const results = [], seen = new Set();

    const candidates = [
        ...document.querySelectorAll('li[class*="MuiListItem"]'),
        ...document.querySelectorAll('div[role="listitem"]'),
        ...document.querySelectorAll('div[class*="MuiBox"]'),
    ];

    for (const d of candidates) {
        const r = d.getBoundingClientRect();
        const style = window.getComputedStyle(d);

        if (r.width <= 80 || r.width >= halfW) continue;
        if (r.height <= 40 || r.height >= 350)  continue;
        if (style.cursor !== 'pointer')          continue;

        const innerText = (d.innerText || '').trim();
        if (!innerText || innerText.length < 5) continue;

        const lines = innerText.split('\n').map(l => l.trim()).filter(Boolean);
        if (!lines.length) continue;

        const title = lines[0];
        if (JUNK.has(title.toLowerCase())) continue;
        if (title.length < 4) continue;
        if (seen.has(title)) continue;
        seen.add(title);

        const chipEls = d.querySelectorAll('span[class*="MuiChip-label"]');
        const chips = Array.from(chipEls).map(c => (c.innerText || '').trim()).filter(Boolean);

        const chipSet = new Set(chips.map(c => c.toLowerCase()));
        let uniOnCard = '';
        for (let i = 1; i < lines.length; i++) {
            const l = lines[i];
            if (!chipSet.has(l.toLowerCase()) && l.length > 3 && l.length < 100 &&
                !l.toLowerCase().includes('tuition') && !l.toLowerCase().includes('fee')) {
                uniOnCard = l;
                break;
            }
        }

        const tuition = lines.find(l =>
            l.toLowerCase().includes('fee') || l.toLowerCase().includes('tuition')
        ) || '';

        results.push({ title, chips, uniOnCard, tuition });
    }
    return results;
}
"""

# ══════════════════════════════════════════════════════════════
# JAVASCRIPT  — Scroll the left panel
# ══════════════════════════════════════════════════════════════
SCROLL_PANEL_JS = r"""
(step) => {
    const halfW = window.innerWidth * 0.65;
    for (const d of document.querySelectorAll('div')) {
        const r = d.getBoundingClientRect();
        if (r.width > 80 && r.width < halfW &&
            r.height > window.innerHeight * 0.4 &&
            d.scrollHeight > d.clientHeight + 50) {
            d.scrollTop += step;
            return {
                scrollTop:    d.scrollTop,
                scrollHeight: d.scrollHeight,
                clientHeight: d.clientHeight,
                atBottom:     d.scrollTop + d.clientHeight >= d.scrollHeight - 5
            };
        }
    }
    window.scrollBy(0, step);
    return {
        scrollTop:    window.scrollY,
        scrollHeight: document.body.scrollHeight,
        clientHeight: window.innerHeight,
        atBottom:     window.scrollY + window.innerHeight >= document.body.scrollHeight - 5
    };
}
"""

# ══════════════════════════════════════════════════════════════
# JAVASCRIPT  — Click a card by title
# ══════════════════════════════════════════════════════════════
CLICK_CARD_JS = r"""
(title) => {
    const halfW = window.innerWidth * 0.65;
    const candidates = [
        ...document.querySelectorAll('li[class*="MuiListItem"]'),
        ...document.querySelectorAll('div[role="listitem"]'),
        ...document.querySelectorAll('div[class*="MuiBox"]'),
    ];
    for (const d of candidates) {
        const r = d.getBoundingClientRect();
        const style = window.getComputedStyle(d);
        if (r.width > 80 && r.width < halfW &&
            r.height > 40 && r.height < 350 &&
            style.cursor === 'pointer') {
            const firstLine = (d.innerText || '').trim().split('\n')[0].trim();
            if (firstLine === title) {
                d.scrollIntoView({ behavior: 'instant', block: 'center' });
                d.click();
                return true;
            }
        }
    }
    return false;
}
"""

# ══════════════════════════════════════════════════════════════
# JAVASCRIPT  — Extract detail panel data (GLOBAL SCOPE & SIBLINGS)
# ══════════════════════════════════════════════════════════════
EXTRACT_DETAIL_JS = r"""
() => {
    // Grab all text globally so the AI gets the full page context including the drawer
    const allText = document.body.innerText || '';

    // ── University Name ──
    let uniName = '';
    const nacpbgEl = document.querySelector('span[class*="nacpbg"]');
    if (nacpbgEl) uniName = (nacpbgEl.innerText || '').trim();

    if (!uniName) {
        for (const sp of document.querySelectorAll('span[class*="bodySmStrong"], span[class*="bodySm"]')) {
            const t = (sp.innerText || '').trim();
            const tl = t.toLowerCase();
            if (t.length > 4 && t.length < 100 && !t.includes('\n') &&
                !/^\d/.test(t) && !tl.includes('deadline') && !tl.includes('tuition') &&
                !tl.includes('semester') && !tl.includes('none') && !tl.includes('login') && 
                !tl.includes('apply')) {
                uniName = t;
                break;
            }
        }
    }

    // ── Deadlines and Tuition (Precise Sibling Target Globally) ──
    let deadlineSummer = '', deadlineWinter = '', tuitionFee = '';

    const allSpans = Array.from(document.querySelectorAll('span'));
    for (let i = 0; i < allSpans.length; i++) {
        const span = allSpans[i];
        const labelText = (span.innerText || '').trim().toLowerCase();

        const isSummerDL  = labelText.includes('summer') && labelText.includes('deadline');
        const isWinterDL  = labelText.includes('winter') && labelText.includes('deadline');
        const isTuition   = labelText === 'tuition fee' || labelText === 'tuition';

        if (!isSummerDL && !isWinterDL && !isTuition) continue;

        // The exact DOM structure shows the value is the next immediate element sibling
        const nextSibling = span.nextElementSibling;
        if (nextSibling) {
            const value = (nextSibling.innerText || '').trim();
            if (value && value.length < 50) {
                if (isSummerDL && !deadlineSummer) deadlineSummer = value;
                if (isWinterDL && !deadlineWinter) deadlineWinter = value;
                if (isTuition  && !tuitionFee)     tuitionFee     = value;
            }
        }
    }

    // ── Chips ──
    const chips = Array.from(document.querySelectorAll('span[class*="MuiChip-label"]'))
                       .map(e => (e.innerText || '').trim())
                       .filter(Boolean);

    return {
        allText,
        uniName,
        chips,
        deadlineSummer,
        deadlineWinter,
        tuitionFee,
    };
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
    for old, new in {
        "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
        "\u2013": "-", "\u2014": "-", "\u2026": "...", "\u00a0": " "
    }.items():
        s = s.replace(old, new)
    return re.sub(r"\s+", " ", s).strip()

def clean_degree(s: str) -> str:
    parts = [p.strip() for p in re.split(r"[,;/]+", s) if p.strip()]
    out, seen = [], set()
    for p in parts:
        p = re.sub(r"^Master'?s?(\s+of\s+\w+)?$", "Masters", p, flags=re.IGNORECASE)
        p = re.sub(r"^Bachelor'?s?(\s+of\s+\w+)?$", "Bachelors", p, flags=re.IGNORECASE)
        if p.lower() not in seen:
            seen.add(p.lower())
            out.append(p)
    return ", ".join(out)

def parse_deadline(s: str) -> date | None:
    s = clean(s).strip()
    if not s or s.lower() in ("none", "n/a", "-", ""):
        return None
    for fmt in [
        "%Y-%m-%d", "%d %B %Y", "%d %b %Y",
        "%B %d, %Y", "%b %d, %Y",
        "%d/%m/%Y",  "%m/%d/%Y",
        "%d-%B-%Y",  "%d %B, %Y",
    ]:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            pass
    m = re.search(
        r"(\d{1,2})\s+"
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})",
        s, re.I
    )
    if m:
        try:
            return datetime.strptime(
                f"{m.group(1)} {m.group(2).title()} {m.group(3)}", "%d %B %Y"
            ).date()
        except ValueError:
            pass
    return None

def best_deadline(summer: str, winter: str) -> str:
    s_clean = clean(summer)
    w_clean = clean(winter)

    s_absent = not s_clean or s_clean.lower() in ("none", "-", "n/a")
    w_absent = not w_clean or w_clean.lower() in ("none", "-", "n/a")

    if s_absent and w_absent:
        return ""
    if s_absent:
        parsed = parse_deadline(w_clean)
        return parsed.isoformat() if parsed else w_clean
    if w_absent:
        parsed = parse_deadline(s_clean)
        return parsed.isoformat() if parsed else s_clean

    ps = parse_deadline(s_clean)
    pw = parse_deadline(w_clean)

    if ps and pw:
        future_s = ps >= TODAY
        future_w = pw >= TODAY
        if future_s and future_w:
            chosen = min(ps, pw)
        elif future_s:
            chosen = ps
        elif future_w:
            chosen = pw
        else:
            chosen = max(ps, pw) 
        return chosen.isoformat()

    if ps: return ps.isoformat()
    if pw: return pw.isoformat()

    return w_clean or s_clean


# ══════════════════════════════════════════════════════════════
# AI EXTRACTION
# ══════════════════════════════════════════════════════════════

def build_prompt(text, name, chips):
    return (
        "You are a precise data extraction engine.\n\n"
        "STRICT RULES:\n"
        "- Return ONLY valid JSON, no preamble, no markdown fences.\n"
        "- Plain ASCII strings only.\n"
        "- Write Masters (not Master's), Bachelors (not Bachelor's).\n"
        "- degree_type: comma-separated values from: "
        "  Bachelors, Masters, MBA, PhD, Executive MBA, Certificate.\n"
        "- scholarship_type: Extract the tuition fee info. If free/none, output 'No Tuition Fee'. If there is a fee, output 'Tuition Required' or the fee amount.\n"
        "- deadline: Extract the application deadline in YYYY-MM-DD format (e.g., 2024-07-15). If multiple, pick the upcoming one. If none found, leave empty string.\n"
        "- host_country: where the student physically studies. Default Germany.\n"
        "- field_of_study: comma-separated academic subject fields.\n\n"
        '{"scholarship_type":"","deadline":"","host_country":"","degree_type":"","field_of_study":""}\n\n'
        f"Program: {name}\n"
        f"Chips/Tags: {', '.join(chips) if chips else 'N/A'}\n\n"
        f"Text:\n{text[:7000]}"
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

async def ai_extract_one(text: str, name: str, chips: list, sem: asyncio.Semaphore) -> dict:
    async with sem:
        try:
            resp = await ai_client.chat.completions.create(
                model="gpt-4o-mini",
                temperature=0,
                messages=[
                    {"role": "system", "content": "Structured data extraction. Return only valid JSON."},
                    {"role": "user",   "content": build_prompt(text, name, chips)},
                ],
                max_tokens=400,
            )
            raw = resp.choices[0].message.content or ""
        except Exception as e:
            print(f"  ⚠  OpenAI [{name[:30]}]: {e}")
            return dict(EMPTY)
    data = parse_json_safe(raw)
    out = dict(EMPTY)
    for k in out:
        if k in data and data[k] is not None:
            out[k] = clean(str(data[k]))
    return out


# ══════════════════════════════════════════════════════════════
# SCRAPE ONE CARD
# ══════════════════════════════════════════════════════════════

async def scrape_card(page, card: dict, seq: int, total: int, ai_sem: asyncio.Semaphore):
    name        = clean(card.get("title", ""))
    chips_card  = card.get("chips", [])
    uni_on_card = clean(card.get("uniOnCard", ""))
    tuition_fb  = clean(card.get("tuition", ""))

    if not name or is_junk_title(name):
        return None

    clicked = await page.evaluate(CLICK_CARD_JS, name)
    if not clicked:
        try:
            await page.get_by_text(name, exact=True).first.click(timeout=3000)
            clicked = True
        except Exception:
            pass
    if not clicked:
        print(f"  ⚠  ({seq}/{total}) Click failed: {name[:50]}")
        return None

    await page.wait_for_timeout(DETAIL_WAIT)

    detail = await page.evaluate(EXTRACT_DETAIL_JS)

    text           = clean(detail.get("allText", ""))
    chips_detail   = detail.get("chips", [])
    uni_name       = clean(detail.get("uniName", "")) or uni_on_card
    deadline_sum   = clean(detail.get("deadlineSummer", ""))
    deadline_win   = clean(detail.get("deadlineWinter", ""))
    dom_tuition    = clean(detail.get("tuitionFee", "")) or tuition_fb

    all_chips = list(dict.fromkeys(chips_card + chips_detail))

    if len(text) < 50:
        print(f"  ⚠  ({seq}/{total}) Empty panel: {name[:50]}")
        return None

    ai_data = await ai_extract_one(text, name, all_chips, ai_sem)

    # 1. Resolve Deadline (DOM first, then AI)
    deadline_out = best_deadline(deadline_sum, deadline_win)
    if not deadline_out:
        ai_dl = clean(ai_data.get("deadline", ""))
        if ai_dl and ai_dl.lower() not in ("none", "", "n/a"):
            parsed = parse_deadline(ai_dl)
            deadline_out = parsed.isoformat() if parsed else ai_dl

    # 2. Resolve Scholarship Type / Tuition (DOM first, then AI, then card Fallback)
    scholarship_out = dom_tuition
    if not scholarship_out or scholarship_out.lower() in ("none", "n/a", ""):
        scholarship_out = clean(ai_data.get("scholarship_type", ""))
    if not scholarship_out:
        scholarship_out = tuition_fb

    # Normalize free tuition language
    if scholarship_out.lower() in ("none", "free", "0", "no tuition", "no tuition fee"):
        scholarship_out = "No Tuition Fee"

    # 3. Resolve Degree
    degree_raw = clean(ai_data.get("degree_type", ""))
    if not degree_raw:
        degree_chips = [
            c for c in all_chips
            if re.search(r"\b(master|bachelor|phd|mba|doctorate|b\.eng|m\.sc|m\.a\.|b\.sc)\b", c, re.I)
        ]
        degree_raw = ", ".join(degree_chips)
    degree_out = clean_degree(degree_raw)

    print(
        f"  ✅  ({seq}/{total})  {name[:45]:<45}  "
        f"uni=[{uni_name[:30] or 'N/A'}]  "
        f"dl=[{deadline_out or 'N/A'}]  "
        f"tuition=[{scholarship_out or 'N/A'}]"
    )

    return {
        "Scholarship Name": name,
        "Scholarship Link": "N/A (SPA)",
        "Official Website": uni_name,
        "Scholarship Type": scholarship_out,
        "Deadline":         deadline_out,
        "Host Country":     clean(ai_data.get("host_country", "")) or "Germany",
        "Degree Type":      degree_out,
        "Field of Study":   clean(ai_data.get("field_of_study", "")),
    }


# ══════════════════════════════════════════════════════════════
# BATCH SCRAPE
# ══════════════════════════════════════════════════════════════

async def scrape_batch(page, new_cards, scraped, ai_sem, global_seq, global_tot):
    records = []
    for i, card in enumerate(new_cards):
        name = clean(card.get("title", ""))
        if is_junk_title(name):
            scraped.add(name)
            continue
        try:
            res = await scrape_card(page, card, global_seq + i, global_tot, ai_sem)
            if isinstance(res, dict):
                records.append(res)
        except Exception as e:
            print(f"  ⚠  Error [{name[:40]}]: {e}")
        scraped.add(name)
    return records


# ══════════════════════════════════════════════════════════════
# MAIN
# ══════════════════════════════════════════════════════════════

async def main():
    print("═" * 64)
    print("  Expatrio Study-Finder Scraper  →  Google Sheets  [FINAL FIXED]")
    print(f"  Tab : {SHEET_TAB}")
    print("═" * 64)
    print("\n  🔗  Connecting to Google Sheets ...")
    existing = load_existing_names()
    print(f"  🗃   {len(existing)} already-saved programs loaded\n")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(
            headless=True,
            args=[
                "--no-sandbox", "--disable-setuid-sandbox",
                "--disable-dev-shm-usage", "--disable-gpu",
            ]
        )
        ctx = await browser.new_context(
            viewport={"width": 1400, "height": 900},
            device_scale_factor=1,
            is_mobile=False,
            has_touch=False,
            extra_http_headers={
                "User-Agent": (
                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                    "Chrome/122.0.0.0 Safari/537.36"
                ),
                "Accept-Language": "en-US,en;q=0.9",
            },
        )
        page = await ctx.new_page()
        page.set_default_navigation_timeout(NAV_TIMEOUT)

        print("🌐  Loading page ...")
        await page.goto(START_URL, wait_until="domcontentloaded", timeout=NAV_TIMEOUT)
        print("  ⏳  Waiting 6s for React to render ...")
        await page.wait_for_timeout(6000)

        ai_sem = asyncio.Semaphore(MAX_AI_CONCURRENT)
        scraped = set(existing)
        saved = skipped = round_n = no_new_streak = 0

        print("\n🔄  Starting scrape → scroll → scrape loop ...\n")

        while True:
            round_n += 1
            all_cards = await page.evaluate(COLLECT_CARDS_JS)

            all_cards = [c for c in all_cards if not is_junk_title(c.get("title", ""))]
            new_cards  = [c for c in all_cards if clean(c.get("title", "")) not in scraped]

            if new_cards:
                no_new_streak = 0
                print(f"\n  📋  Round {round_n} — {len(new_cards)} new cards  (seen: {len(scraped)})")
                records = await scrape_batch(
                    page, new_cards, scraped, ai_sem,
                    saved + 1, saved + len(new_cards)
                )
                if records:
                    flush_to_sheets(records)
                    saved   += len(records)
                    skipped += len(new_cards) - len(records)
                    print(f"\n  💾  Saved {len(records)} to Sheets  [total={saved}]\n")
            else:
                no_new_streak += 1
                print(f"  ↕   Round {round_n} — no new cards ({no_new_streak}/5)")
                if no_new_streak >= 5:
                    print("\n  ✅  Done — no new cards for 5 rounds!")
                    break

            scroll_info = await page.evaluate(SCROLL_PANEL_JS, SCROLL_STEP)
            await page.wait_for_timeout(int(SCROLL_PAUSE * 1000))

            if scroll_info.get("atBottom"):
                all_cards  = await page.evaluate(COLLECT_CARDS_JS)
                all_cards  = [c for c in all_cards if not is_junk_title(c.get("title", ""))]
                final_new  = [c for c in all_cards if clean(c.get("title", "")) not in scraped]
                if final_new:
                    records = await scrape_batch(
                        page, final_new, scraped, ai_sem,
                        saved + 1, saved + len(final_new)
                    )
                    if records:
                        flush_to_sheets(records)
                        saved += len(records)
                print("\n  ✅  Bottom of list reached — done!")
                break

        await browser.close()

    print("\n" + "═" * 64)
    print(f"  ✅  DONE  |  Saved: {saved}  |  Skipped: {skipped}")
    print(f"  Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}")
    print("═" * 64)


if __name__ == "__main__":
    asyncio.run(main())