<div align="center">

# 🎓 Scholarships-Scraper

### AI-Powered Scholarship Intelligence Pipeline

**An automated, high-throughput scraping suite that discovers, structures, and aggregates scholarship opportunities from 12+ academic portals — powered by Playwright and OpenAI.**
---

## 📌 Overview

**Scholarships-Scraper** is a production-grade data-mining suite built to solve a deceptively hard problem: scholarship listings are scattered across dozens of websites, each with its own layout, pagination scheme, and writing style.

Rather than relying on brittle, site-specific regex or XPath rules to interpret human-written descriptions, this project uses a **hybrid architecture** — deterministic browser automation handles navigation, while **Large Language Models (OpenAI)** handle the unpredictable part: reading raw scholarship text and converting it into clean, structured JSON. The result is centralized into a single, continuously updated **Google Sheet**.

> 12 independent scrapers. One unified pipeline. Zero manual data entry.

---

## ✨ Key Features

- 🔍 **Multi-Source Discovery** — Crawls 12 major scholarship portals, each with custom handling for pagination, infinite scroll, and "Load More" interactions.
- 🧠 **AI-Powered Structuring** — Uses OpenAI to transform messy, free-text scholarship descriptions into strict, standardized JSON fields.
- ⚡ **Performance-First Scraping** — Blocks images, fonts, and CSS during navigation; strips DOM noise (headers, footers, navbars) before parsing.
- 📅 **Smart Date Normalization** — A custom parser standardizes inconsistent date formats (`"Mid-February 2025"`, `"15/02/2025"`, etc.) into ISO 8601, automatically dropping expired listings.
- 🧩 **Deduplication Built-In** — Cross-checks newly discovered URLs against existing Google Sheet records to avoid duplicate entries.
- 📊 **Centralized Data Output** — All scrapers funnel into a single Google Sheet via `gspread`, with rate-limit-safe batched writes.
- 🛠️ **Hybrid Execution Strategies** — Mixes `async Playwright`, `sync Playwright`, `requests + BeautifulSoup`, and pure DOM parsing depending on what each target site actually needs (no one-size-fits-all overhead).
- 🔐 **Security-Conscious by Design** — Credentials and API keys are isolated via `.env` and `credentials.json`, both excluded from version control.

---

## 🏗️ Architecture: The 4-Stage Pipeline

Every scraper in this project follows the same disciplined pipeline, regardless of the target site's complexity:

```
┌──────────────────┐     ┌──────────────────┐     ┌──────────────────┐     ┌──────────────────┐
│   STAGE 1         │     │   STAGE 2         │     │   STAGE 3         │     │   STAGE 4         │
│   Discovery       │ ──▶ │   Extraction      │ ──▶ │   AI Structuring  │ ──▶ │   Persistence     │
│                    │     │                    │     │                    │     │                    │
│ Listing pages,     │     │ Resource blocking, │     │ Raw text → OpenAI  │     │ Batched writes to  │
│ pagination, scroll,│     │ DOM stripping,     │     │ → strict JSON      │     │ Google Sheets via  │
│ deduplication      │     │ core text isolation│     │ schema + date norm │     │ gspread            │
└──────────────────┘     └──────────────────┘     └──────────────────┘     └──────────────────┘
```

### Stage 1 — Discovery
Navigates listing pages and handles whatever discovery mechanism the site uses: URL-based pagination, infinite scroll, or "Load More" buttons — while filtering out scholarships already present in the destination sheet.

### Stage 2 — Detail Extraction
Visits each newly discovered scholarship page with images, fonts, and CSS blocked for speed. Injected JavaScript strips away navigational chrome, isolating the core article text.

### Stage 3 — AI Structuring
The cleaned text is passed to the OpenAI API under a strict system prompt, returning pure JSON across a fixed schema: `scholarship_type`, `deadline`, `host_country`, `degree_type`, `field_of_study`. A custom parser then normalizes all dates and discards anything past its deadline.

### Stage 4 — Persistence
Structured records are appended to a shared Google Sheet using a Service Account, with writes batched (5–12 records at a time) to stay safely under API rate limits.

---

## 🌐 Supported Scholarship Sources

| Scraper | Target Website | Execution Strategy | Notable Technique |
|---|---|---|---|
| `365.py` | Scholarships365 | Async Playwright + OpenAI | Async pagination with summary-card pre-extraction |
| `bright.py` | BrightScholarship | Async Playwright + OpenAI | Semaphore-based concurrency for high-volume listings |
| `corner.py` | ScholarshipsCorner | Async Playwright + OpenAI | Strict text-line evaluation to preserve context |
| `Opp.py` | OpportunitiesCorners | Async Playwright + OpenAI | Custom `clean_degree()` normalization logic |
| `top.py` | TopUniversities | Async Playwright + OpenAI | Explicit wait conditions for React/Vue shadow DOMs |
| `idp.py` | IDP Find-a-Scholarship | Async Playwright + OpenAI | 3-hop traversal: Listing → University Page → Detail |
| `exp.py` | Expatrio | Async Playwright + OpenAI | SPA-aware routing without full page reloads |
| `wms.py` | WeMakeScholars | Async Playwright + OpenAI | Recursive "Load More" automation + Provider Name field |
| `abro.py` | StudyAbroad | Sync Playwright + OpenAI | BeautifulSoup + Playwright hybrid, optional Excel export |
| `dev.py` | Scholars4Dev | Requests + BS4 + OpenAI | Pure HTTP scraping — no browser overhead |
| `roar.py` | ScholarshipRoar | Requests + Playwright + OpenAI | Requests for listings, Playwright for dynamic detail pages |
| `Daad.py` | DAAD (Germany) | Sync Playwright (No AI) | Pure DOM parsing — structured HTML, no LLM cost needed |

---

## 🛠️ Tech Stack

| Layer | Technology |
|---|---|
| Browser Automation | Playwright (sync & async) |
| Static Scraping | `requests`, `BeautifulSoup` |
| AI Structuring | OpenAI API |
| Data Storage | Google Sheets via `gspread` |
| Config Management | `python-dotenv` |
| Optional Export | `pandas`, `openpyxl` |

---

## 🚀 Getting Started

### Prerequisites
- Python 3.10+
- A Google Cloud Service Account with Sheets API access
- An OpenAI API key

### Installation

```bash
git clone https://github.com/<your-username>/scholarships-scraper.git
cd scholarships-scraper
pip install -r requirements.txt
playwright install
```

### Configuration

1. Place your Google Cloud Service Account key as `credentials.json` in the project root.
2. Create a `.env` file:

```env
OPENAI_API_KEY=your_openai_api_key_here
GOOGLE_SHEET_ID=your_target_sheet_id_here
```

### Running a Scraper

```bash
python 365.py
# or any other scraper, e.g.
python Daad.py
```

---

## 🔐 Security

- `.env` and `credentials.json` are **never** committed — both are excluded via `.gitignore`.
- API keys are loaded at runtime only, never hardcoded.
- `requirements.txt` pins dependencies to avoid environment drift between machines.

---

## 🗺️ Roadmap

- [ ] Unified CLI to trigger all scrapers from a single entry point
- [ ] Scheduling support (cron / GitHub Actions) for fully autonomous daily runs
- [ ] Web dashboard for browsing aggregated scholarship data
- [ ] Email/Telegram alerts for new high-relevance scholarships

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome. Feel free to check the [issues page](../../issues) or open a pull request.

---

## 📄 License

This project is licensed under the MIT License — see the [LICENSE](LICENSE) file for details.

---

<div align="center">

**Built for students who deserve to find opportunities without the endless tab-switching.**

</div>
