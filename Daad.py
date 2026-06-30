#!/usr/bin/env python3
import time
import gspread
from google.oauth2.service_account import Credentials
from playwright.sync_api import sync_playwright

# ==================================================
# CONFIG
# ==================================================
START_URL = "https://www.daad.pk/en/"

# ==================================================
# GOOGLE SHEETS CONFIG
# ==================================================
CREDENTIALS_FILE = "credentials.json"
SHEET_ID         = "1Me3gRF_0gnCbn1p3c9hCvS77CTEkuzStaqd4AbiQ6kw"
SHEET_TAB        = "daad"

# DAAD has completely different columns — social feed posts
COLUMN_HEADERS = [
    "post_title", "author", "date", "full_text",
    "facebook_post_link", "image_url", "official_link",
]

SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
]

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
            # Write header IMMEDIATELY on tab creation
            _gs_worksheet.append_row(COLUMN_HEADERS, value_input_option="USER_ENTERED")
            _header_written = True
            print(f"  📋  Created tab '{SHEET_TAB}' with headers")
    return _gs_worksheet


def load_existing_links() -> set:
    """Return set of already-saved facebook_post_link + official_link values."""
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
        links   = set()
        for col_name in ("facebook_post_link", "official_link"):
            try:
                idx = headers.index(col_name)
                for i in range(1, len(rows)):
                    if len(rows[i]) > idx and rows[i][idx].strip():
                        links.add(rows[i][idx].strip())
            except ValueError:
                pass
        return links
    except Exception as e:
        print(f"  ⚠  Could not load existing links: {e}")
        return set()


def save_to_sheets(records: list[dict], retries: int = 3):
    """Append new records to Google Sheets."""
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
            print(f"  ⚠  save_to_sheets error: {e}")
            time.sleep(5)


# ==================================================
# LOAD EXISTING LINKS
# ==================================================
print("🌐 Connecting to Google Sheets ...")
existing_links = load_existing_links()
print(f"ℹ️  {len(existing_links)} existing links loaded from '{SHEET_TAB}'\n")

# ==================================================
# SCRAPER
# ==================================================
results = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    page = browser.new_page()
    page.goto(START_URL, wait_until="domcontentloaded", timeout=60000)

    page.evaluate("window.scrollTo(0, document.body.scrollHeight)")
    time.sleep(2)

    try:
        page.wait_for_selector(".c-social-feed__item", timeout=30000)
    except Exception:
        pass

    data = page.evaluate(
        """
        () => {
            const items = Array.from(document.querySelectorAll('.c-social-feed__item'));
            return items.map(item => {
                const dds    = Array.from(item.querySelectorAll('dl dd'));
                const author = dds[0]?.innerText.trim() || '';
                const date   = dds[1]?.innerText.trim() || '';

                const contentEl = item.querySelector('.c-social-feed__item-content');
                const fullText  = contentEl?.innerText.trim() || '';

                let title = '';
                const titleEl = item.querySelector('.c-slide-teaser__item-title, .c-social-feed__item-title');
                if (titleEl) {
                    title = titleEl.innerText.trim();
                } else if (fullText) {
                    title = fullText.split('\\n')[0].trim();
                }

                const fbLink   = item.querySelector('a.o-more-link')?.getAttribute('href') || '';
                const img      = item.querySelector('.c-slide-teaser__asset-social img');
                const imageUrl = img?.getAttribute('src') || img?.getAttribute('data-src') || '';

                let officialLink = '';
                const anchors = Array.from(item.querySelectorAll('.c-social-feed__item-content a'));
                for (const a of anchors) {
                    const href       = a.getAttribute('href') || '';
                    const isFacebook = href.includes('facebook.com') || href.includes('fb.com');
                    if (href && !isFacebook) { officialLink = href; break; }
                }

                return {
                    post_title:          title,
                    author,
                    date,
                    full_text:           fullText,
                    facebook_post_link:  fbLink,
                    image_url:           imageUrl,
                    official_link:       officialLink
                };
            });
        }
        """
    )

    if isinstance(data, list):
        results.extend(data)

    browser.close()

# ==================================================
# FILTER AGAINST EXISTING ENTRIES
# ==================================================
filtered = []
skipped  = 0

for r in results:
    fb  = (r.get("facebook_post_link") or "").strip()
    off = (r.get("official_link")      or "").strip()
    if (fb and fb in existing_links) or (off and off in existing_links):
        skipped += 1
        print(f"⛔ Skipping already-present: fb='{fb}'  off='{off}'")
    else:
        filtered.append(r)

print(f"\n➡  Collected {len(results)} items  |  {len(filtered)} new  |  {skipped} skipped\n")

# ==================================================
# SAVE TO GOOGLE SHEETS
# ==================================================
if filtered:
    save_to_sheets(filtered)
    print(f"✅  DONE — saved {len(filtered)} records to Sheets (tab: '{SHEET_TAB}')")
    print(f"🔗  Sheet: https://docs.google.com/spreadsheets/d/{SHEET_ID}")
else:
    print("⚠  No new records to write.")