#!/usr/bin/env python3
import os
import json
import time
import re
import requests
from datetime import datetime, date
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from openai import OpenAI
from dotenv import load_dotenv

import gspread
from google.oauth2.service_account import Credentials

load_dotenv()

# ==================================================
# CONFIG
# ==================================================
BASE_URL = "https://scholarshiproar.com/category/scholarships/"
START_PAGE = 1
END_PAGE = 18
BATCH_SIZE = 5

OPENAI_API_KEY = os.environ["OPENAI_API_KEY"]

TODAY = date.today()

# ==================================================
# GOOGLE SHEETS CONFIG
# ==================================================
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "scholarships_Roar"

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
# FILTERS
# ==================================================
REQUIRED_KEYWORDS = [
    "scholarship", "fellowship", "grant", "award", "funded",
    "intern", "internship", "program", "programme"
]

BLOCKED_KEYWORDS = [
    "motivation letter", "statement of purpose", "sop",
    "how to", "tips", "template", "examples", "essay",
    "cv", "resume"
]

def is_valid_scholarship(title: str, categories: list[str]) -> bool:
    t = title.lower()
    cats = " ".join(categories).lower()

    if not any(k in t or k in cats for k in REQUIRED_KEYWORDS):
        return False

    if any(k in t or k in cats for k in BLOCKED_KEYWORDS):
        return False

    return True

# ==================================================
# DEADLINE PARSING / STANDARDIZATION
# ==================================================
def parse_deadline(deadline_str: str) -> date | None:
    if not deadline_str or not deadline_str.strip():
        return None

    s = deadline_str.strip()
    formats = [
        "%d %B %Y", "%d %b %Y", "%B %d, %Y", "%b %d, %Y",
        "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%y",
        "%d-%b-%Y", "%d-%B-%Y", "%d %B, %Y", "%d %b, %Y",
        "%b %d %Y", "%B %d %Y", "%d/%m/%y", "%m/%d/%y"
    ]

    for fmt in formats:
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue

    m = re.search(
        r"(\d{1,2})\s+(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{4})",
        s, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(f"{m.group(1)} {m.group(2)} {m.group(3)}", "%d %B %Y").date()
        except ValueError:
            pass

    m = re.search(
        r"(January|February|March|April|May|June|July|August|September|October|November|December)\s+(\d{1,2}),?\s+(\d{4})",
        s, re.IGNORECASE
    )
    if m:
        try:
            return datetime.strptime(f"{m.group(2)} {m.group(1)} {m.group(3)}", "%d %B %Y").date()
        except ValueError:
            pass

    m = re.search(r"\b(\d{4}-\d{2}-\d{2})\b", s)
    if m:
        try:
            return datetime.strptime(m.group(1), "%Y-%m-%d").date()
        except ValueError:
            pass

    m = re.search(r"\b(\d{1,2}/\d{1,2}/\d{2,4})\b", s)
    if m:
        try:
            t = m.group(1)
            try:
                return datetime.strptime(t, "%d/%m/%Y").date()
            except ValueError:
                return datetime.strptime(t, "%m/%d/%Y").date()
        except Exception:
            pass

    return None

def standardize_deadline_iso_or_raw(deadline_str: str) -> str:
    if not deadline_str:
        return ""
    parsed = parse_deadline(deadline_str)
    return parsed.isoformat() if parsed else deadline_str.strip()

# ==================================================
# STEP 1 — DISCOVER SCHOLARSHIP LINKS
# ==================================================
headers = {"User-Agent": "Mozilla/5.0"}
scholarships = []

def get_page_url(page_num: int) -> str:
    if page_num == 1:
        return BASE_URL
    return f"{BASE_URL}page/{page_num}/"

print("🌐 Fetching pages...")

for page_num in range(START_PAGE, END_PAGE + 1):
    page_url = get_page_url(page_num)
    print(f"➡ Page {page_num}: {page_url}")

    try:
        resp = requests.get(page_url, headers=headers, timeout=30)
        resp.raise_for_status()
        html = resp.text
    except Exception as e:
        print(f"⚠ Failed to fetch page {page_num}: {e}")
        continue

    soup = BeautifulSoup(html, "html.parser")

    for article in soup.select("article.post"):
        title_el = article.select_one("h2.entry-title a")
        if not title_el:
            continue

        title = title_el.get_text(strip=True)
        link = title_el.get("href", "").strip()

        cat_els = article.select(".cat-links a")
        categories = [c.get_text(strip=True) for c in cat_els]

        if title and link and is_valid_scholarship(title, categories):
            scholarships.append({"name": title, "link": link})
            print(f"✔ {title}")
        else:
            print(f"⛔ Skipped: {title}")

print(f"\n📦 Found {len(scholarships)} scholarships\n")

# ==================================================
# CHECK SPREADSHEET FOR EXISTING LINKS (skip if present)
# ==================================================
print("🔗 Connecting to Google Sheets...")
existing_links = load_existing_links()

filtered_scholarships = []
skipped_count = 0
for s in scholarships:
    if s["link"] in existing_links:
        skipped_count += 1
        print(f"⛔ Skipping already-present link in sheet: {s['link']}")
    else:
        filtered_scholarships.append(s)

scholarships = filtered_scholarships

print(f"\n➡ Will scrape {len(scholarships)} links; skipped {skipped_count} already-present links from Google Sheets (tab: {SHEET_TAB})\n")

# ==================================================
# STEP 2 — SCRAPE + AI STRUCTURE (with batch writes)
# ==================================================
print("🧭 Launching browser...")

results_batch = []
success_count = 0

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()

    for s in scholarships:
        print(f"\n🔍 Scraping: {s['name']}")

        try:
            page.goto(s["link"], wait_until="domcontentloaded", timeout=60000)
        except Exception as e:
            print("⚠ Page load failed:", e)
            continue

        try:
            page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
            time.sleep(1)
        except:
            pass

        text = page.evaluate("""
        () => {
            const main = document.querySelector('.entry-content') || document.body;
            return main ? main.innerText : document.body.innerText;
        }
        """)

        official_website = ""
        try:
            el = page.locator("a:has-text('Official Website')").first
            official_website = el.get_attribute("href") or ""
        except:
            official_website = ""

        prompt = f"""
You are a data extraction engine.

STRICT RULES:
- Return ONLY valid JSON
- No explanations
- No markdown
- No extra text

JSON schema:
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
                model="gpt-4.1",
                input=prompt,
                temperature=0
            )

            chunks = []
            for item in response.output:
                for block in item.content:
                    if hasattr(block, "text"):
                        chunks.append(block.text)

            raw = "".join(chunks).strip()
            data = json.loads(raw)

            if isinstance(data, list):
                data = data[0] if data else {}
            elif not isinstance(data, dict):
                data = {}
        except Exception as e:
            print("⚠ OpenAI failed:", e)
            data = {
                "scholarship_type": "",
                "deadline": "",
                "host_country": "",
                "degree_type": "",
                "field_of_study": ""
            }

        # FIX: guard against deadline being a dict instead of a string
        deadline_raw = data.get("deadline") or ""
        raw_deadline = deadline_raw.strip() if isinstance(deadline_raw, str) else ""
        deadline_value = standardize_deadline_iso_or_raw(raw_deadline)

        record = {
            "Scholarship Name": s["name"],
            "Scholarship Link": s["link"],
            "Official Website": official_website or "",
            "Scholarship Type": data.get("scholarship_type", "") or "",
            "Deadline": deadline_value,
            "Host Country": data.get("host_country", "") or "",
            "Degree Type": data.get("degree_type", "") or "",
            "Field of Study": data.get("field_of_study", "") or ""
        }

        results_batch.append(record)
        success_count += 1
        print("  ✅ Added (deadline ->", (deadline_value or "empty"), ")")

        if success_count % BATCH_SIZE == 0:
            save_batch(results_batch)
            results_batch.clear()

    browser.close()

if results_batch:
    save_batch(results_batch)
    print(f"✅ Wrote final batch of {len(results_batch)} to Google Sheets (tab: {SHEET_TAB})")
else:
    print("\n⚠ No remaining results to write.")

print(f"\n📊 Saved: {success_count} | Skipped (filtered out during discovery / already in sheet): {skipped_count}")
print(f"🔗 Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}")