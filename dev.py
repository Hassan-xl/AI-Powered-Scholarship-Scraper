#!/usr/bin/env python3
import os
import requests
from bs4 import BeautifulSoup
import json
import time
import re
from datetime import datetime, date

import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

# ==================================================
# CONFIG
# ==================================================
BASE_URL = "https://www.scholars4dev.com/tag/scholarships-for-pakistans/"
OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]
BATCH_SIZE = 5
MAX_TEXT_CHARS = 6000
TODAY = date.today()

# Set to True while debugging to print official website and raw/normalized deadlines
VERBOSE = False

# ==================================================
# GOOGLE SHEETS CONFIG
# ==================================================
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "scholars4dev_scholarships"

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
# OPENAI CLIENT
# ==================================================
client = OpenAI(api_key=OPENAI_API_KEY)

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
            print(f"💾 Saved {len(rows_buffer)} records to Google Sheets")
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
# HELPERS
# ==================================================
def trim_text(text, max_chars=MAX_TEXT_CHARS):
    return text[:max_chars]

# ==================================================
# DEADLINE PARSING & NORMALIZATION
# ==================================================
_DATE_FORMATS = [
    "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
    "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%y",
    "%d-%b-%Y", "%d-%B-%Y", "%d %B, %Y", "%d %b, %Y",
    "%b %d %Y", "%B %d %Y", "%d/%m/%y", "%m/%d/%y",
    "%d-%b-%y", "%d-%b-%Y", "%d.%m.%Y",
]

_DATE_REGEXES = [
    r"\d{4}-\d{1,2}-\d{1,2}",
    r"\d{1,2}[/-]\d{1,2}[/-]\d{2,4}",
    r"\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{4}",
    r"(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2},?\s+\d{4}",
    r"\d{1,2}-[A-Za-z]{3}-\d{2,4}",
    r"\d{1,2}\.\d{1,2}\.\d{4}",
]

def try_parse_date(token: str):
    token = token.strip()
    token = token.replace("\u2011", "-")
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(token, fmt).date()
        except Exception:
            continue
    token_no_suffix = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', token, flags=re.IGNORECASE)
    for fmt in _DATE_FORMATS:
        try:
            return datetime.strptime(token_no_suffix, fmt).date()
        except Exception:
            continue
    return None

def extract_dates_from_text(text: str):
    if not text:
        return []

    found_dates = []
    seen = set()

    for rx in _DATE_REGEXES:
        for m in re.finditer(rx, text, flags=re.IGNORECASE):
            token = m.group(0)
            d = try_parse_date(token)
            if d and d not in seen:
                seen.add(d)
                found_dates.append(d)

    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})",
        text, re.IGNORECASE
    )
    if m:
        try:
            d = datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %B %Y").date()
            if d not in seen:
                seen.add(d)
                found_dates.append(d)
        except Exception:
            pass

    m = re.search(
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
        text, re.IGNORECASE
    )
    if m:
        try:
            d = datetime.strptime(f"{m.group(2)} {m.group(1)} {m.group(3)}", "%d %B %Y").date()
            if d not in seen:
                seen.add(d)
                found_dates.append(d)
        except Exception:
            pass

    if not found_dates:
        try:
            d = parse_deadline(text)
            if d:
                found_dates.append(d)
        except Exception:
            pass

    return sorted(found_dates)

def normalize_deadline(raw: str) -> str:
    if not raw:
        return ""

    raw_str = raw.strip()
    raw_str = re.sub(r"(?i)deadline[:\s]*", "", raw_str).strip()
    raw_str = re.sub(r"\s+(?:\(deadline\))", "", raw_str, flags=re.IGNORECASE).strip()

    dates = extract_dates_from_text(raw_str)

    if len(dates) == 1:
        return dates[0].strftime("%Y-%m-%d")
    elif len(dates) >= 2:
        start = dates[0]
        end = dates[-1]
        if start == end:
            return start.strftime("%Y-%m-%d")
        return f"{start.strftime('%Y-%m-%d')} to {end.strftime('%Y-%m-%d')}"
    else:
        return raw_str

def is_deadline_passed(deadline_str: str) -> bool:
    if not deadline_str or not deadline_str.strip():
        return False

    s = deadline_str.strip()
    if "to" in s:
        parts = [p.strip() for p in s.split("to", 1)]
        s = parts[-1]

    try:
        d = datetime.strptime(s, "%Y-%m-%d").date()
        return d < TODAY
    except Exception:
        pass

    d = parse_deadline(s)
    if d is None:
        return False
    return d < TODAY

def parse_deadline(deadline_str: str) -> date | None:
    if not deadline_str or not deadline_str.strip():
        return None

    deadline_str = deadline_str.strip()
    formats = _DATE_FORMATS

    for fmt in formats:
        try:
            return datetime.strptime(deadline_str, fmt).date()
        except ValueError:
            continue

    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{4})",
        deadline_str, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %B %Y").date()
        except ValueError:
            pass

    m = re.search(
        r"(January|February|March|April|May|June|July|August|"
        r"September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
        deadline_str, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(f"{m.group(2)} {m.group(1)} {m.group(3)}", "%d %B %Y").date()
        except ValueError:
            pass

    return None

# ==================================================
# STEP 1 — DISCOVER SCHOLARSHIP LINKS
# ==================================================
print("🌐 Fetching listing pages...")

headers = {"User-Agent": "Mozilla/5.0"}
scholarships = []
page_num = 1

while True:
    url = BASE_URL if page_num == 1 else f"{BASE_URL}page/{page_num}/"
    print(f"📄 Page {page_num}")

    res = requests.get(url, headers=headers, timeout=30)
    if res.status_code != 200:
        break

    soup = BeautifulSoup(res.text, "html.parser")

    cards = soup.select("div.post h2 a")
    if not cards:
        break

    for a in cards:
        title = a.get_text(strip=True)
        link = a.get("href")

        if title and link:
            scholarships.append({"name": title, "link": link})
            print(f"✔ {title}")

    page_num += 1
    time.sleep(0.5)

print(f"\n📦 Found {len(scholarships)} scholarships\n")

# ==================================================
# STEP 1.5 — LOAD EXISTING LINKS FROM GOOGLE SHEETS
# ==================================================
print("🔗 Connecting to Google Sheets...")
existing_links = load_existing_links()
print(f"🔁 Loaded {len(existing_links)} existing links from Google Sheets (tab: {SHEET_TAB})")

# ==================================================
# STEP 2 — SCRAPE + AI STRUCTURE
# ==================================================
print("🧭 Launching browser...")

results_buffer = []
skipped_count = 0
saved_count = 0

def extract_official_website_from_page(page) -> str:
    try:
        href = page.locator("a:has-text('Official Scholarship Website')").first.get_attribute("href") or ""
        if href:
            return href
    except Exception:
        pass

    try:
        xpath_label = "xpath=//p[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'official scholarship website')]/a[1]"
        href = page.locator(xpath_label).first.get_attribute("href") or ""
        if href:
            return href
    except Exception:
        pass

    try:
        xpath_website = "xpath=//p[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'website')]/a[1]"
        href = page.locator(xpath_website).first.get_attribute("href") or ""
        if href:
            return href
    except Exception:
        pass

    try:
        hrefs = page.eval_on_selector_all(".maincontent a", "els => els.map(e => e.href)")
        for h in hrefs:
            if not h:
                continue
            if isinstance(h, str) and h.startswith("http") and "scholars4dev.com" not in h:
                return h
    except Exception:
        pass

    return ""

def _extract_response_text(response):
    # Try direct attribute
    if hasattr(response, "output_text") and response.output_text:
        return response.output_text

    chunks = []

    # response.output might be a list of dicts or objects
    output = getattr(response, "output", None)
    if isinstance(output, list):
        for item in output:
            # item.content may be a list
            content = None
            if isinstance(item, dict):
                content = item.get("content")
            else:
                content = getattr(item, "content", None)

            if isinstance(content, list):
                for block in content:
                    if isinstance(block, dict):
                        # common shape: {"type": "output_text", "text": "..." }
                        if block.get("type") == "output_text" and "text" in block:
                            chunks.append(block["text"])
                        elif "text" in block:
                            chunks.append(block["text"])
                        else:
                            # nested content list
                            inner = block.get("content")
                            if isinstance(inner, list):
                                for b in inner:
                                    if isinstance(b, dict) and "text" in b:
                                        chunks.append(b["text"])
                    else:
                        # object-like block
                        if hasattr(block, "text"):
                            chunks.append(block.text)
            else:
                # item might directly have text
                if isinstance(item, dict) and "text" in item:
                    chunks.append(item["text"])
                elif hasattr(item, "text"):
                    chunks.append(item.text)

    # Fallback: try converting the whole response to string
    if not chunks:
        try:
            s = str(response)
            if s:
                chunks.append(s)
        except Exception:
            pass

    return "\n".join(chunks).strip()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    for idx, s in enumerate(scholarships, start=1):
        # Skip if already in Google Sheets (tab: scholars4dev_scholarships)
        if s["link"] in existing_links:
            print(f"  ⏭ Already in sheet — skipping: {s['link']}")
            skipped_count += 1
            continue

        print(f"\n🔍 Scraping ({idx}/{len(scholarships)}): {s['name']}")

        try:
            page.goto(s["link"], wait_until="domcontentloaded", timeout=60000)
        except Exception:
            print("⚠ Page load failed, skipping")
            continue

        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)
        except Exception:
            pass

        text = page.evaluate("""
        () => {
            const main = document.querySelector('.maincontent');
            return main ? main.innerText : document.body.innerText;
        }
        """)

        text = trim_text(text)

        official_website = ""
        try:
            official_website = extract_official_website_from_page(page) or ""
        except Exception:
            official_website = ""

        # Only print debug info when VERBOSE=True
        if VERBOSE:
            print("  → Official Website:", official_website or "<not found>")

        prompt = f"""
Return ONLY valid JSON. No explanations.

Schema:
{{
  "scholarship_type": "",
  "deadline": "",
  "host_country": "",
  "degree_type": "",
  "field_of_study": ""
}}

Text:
{text}
"""

        try:
            response = client.responses.create(
                model="gpt-4o-mini",
                input=prompt,
                temperature=0
            )

            raw = _extract_response_text(response)

            if VERBOSE:
                print("  → Raw AI response:", raw[:1000])  # print a truncated preview

            # Try to isolate the first JSON object in the output (ignore any surrounding text)
            json_start = raw.find("{")
            json_end = raw.rfind("}")
            json_str = raw
            if json_start != -1 and json_end != -1 and json_end > json_start:
                json_str = raw[json_start:json_end+1]

            # Basic sanitization: remove trailing commas before } or ]
            json_str = re.sub(r",\s*}", "}", json_str)
            json_str = re.sub(r",\s*]", "]", json_str)

            # If still not valid JSON, try swapping single quotes to double quotes (best-effort)
            try:
                data = json.loads(json_str)
            except json.JSONDecodeError:
                alt = json_str.replace("'", '"')
                try:
                    data = json.loads(alt)
                except json.JSONDecodeError:
                    # final attempt: find {...} via regex
                    m = re.search(r"\{(?:[^{}]|(?R))*\}", raw)
                    if m:
                        candidate = m.group(0)
                        candidate = re.sub(r",\s*}", "}", candidate)
                        candidate = re.sub(r",\s*]", "]", candidate)
                        try:
                            data = json.loads(candidate)
                        except Exception:
                            raise ValueError("Empty or invalid JSON")
                    else:
                        raise ValueError("Empty or invalid JSON")

            # If data is not a dict, raise
            if not isinstance(data, dict):
                raise ValueError("Empty or invalid JSON")

        except Exception as e:
            if VERBOSE:
                print("  → AI raw response on failure:", raw if 'raw' in locals() else "<no raw>")
            print("⚠ OpenAI skipped:", e)
            data = {
                "scholarship_type": "",
                "deadline": "",
                "host_country": "",
                "degree_type": "",
                "field_of_study": ""
            }

        raw_deadline_value = data.get("deadline", "") or ""
        normalized_deadline = normalize_deadline(raw_deadline_value)

        if VERBOSE:
            print("  → Raw deadline:", raw_deadline_value or "<empty>")
            print("  → Normalized deadline:", normalized_deadline or "<empty>")

        if is_deadline_passed(normalized_deadline):
            print(f"  ⏭ SKIPPED (deadline passed: {normalized_deadline})")
            skipped_count += 1
            continue

        results_buffer.append({
            "Scholarship Name": s["name"],
            "Scholarship Link": s["link"],
            "Official Website": official_website,
            "Scholarship Type": data.get("scholarship_type", ""),
            "Deadline": normalized_deadline,
            "Host Country": data.get("host_country", ""),
            "Degree Type": data.get("degree_type", ""),
            "Field of Study": data.get("field_of_study", "")
        })

        if len(results_buffer) == BATCH_SIZE:
            save_batch(results_buffer)
            saved_count += len(results_buffer)
            # update existing_links with saved links to avoid re-scraping within this run
            for r in results_buffer:
                ln = r.get("Scholarship Link", "")
                if ln:
                    existing_links.add(str(ln).strip())
            results_buffer.clear()

    browser.close()

if results_buffer:
    save_batch(results_buffer)
    saved_count += len(results_buffer)
    for r in results_buffer:
        ln = r.get("Scholarship Link", "")
        if ln:
            existing_links.add(str(ln).strip())

print(f"\n✅ DONE — Google Sheets update complete")
print(f"📊 Added: {saved_count} | Skipped (expired or existing): {skipped_count}")
print(f"🔗 Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}")