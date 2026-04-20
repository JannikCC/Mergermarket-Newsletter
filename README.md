# Mergermarket Newsletter Automation

Automates the daily Mergermarket intelligence newsletter workflow:
downloads the report, converts it to a formatted Word document with a
Table of Contents, and opens a pre-composed Outlook email for manual review.

## Requirements

- Windows with Microsoft Office (Outlook + Word) installed
- Python 3.10 or later
- A Mergermarket account

## Setup

### 1. Install dependencies

```bat
pip install -r requirements.txt
playwright install chromium
```

### 2. Set Mergermarket credentials

Set these environment variables (only required if you are not already logged
in via saved browser cookies):

```bat
setx MM_USERNAME "your.email@company.com"
setx MM_PASSWORD "your_password"
```

Or set them permanently via **System Properties → Environment Variables**.

## Usage

| Command | Description |
|---|---|
| `python mergermarket_newsletter.py` | Full pipeline for today |
| `python mergermarket_newsletter.py --headless` | Same, browser runs invisibly |
| `python mergermarket_newsletter.py --dry-run C:\path\to\file.xlsx` | Skip download; use existing Excel |
| `python mergermarket_newsletter.py --date 2024-01-08` | Override today's date |
| `python mergermarket_newsletter.py --schedule` | Run as background scheduler (Mon–Fri 08:45) |
| `python mergermarket_newsletter.py --schedule --headless` | Scheduler with invisible browser |

### Monday behaviour

When the run date is a **Monday**, the script automatically sets the search
date range from the **previous Friday** through to today, covering the weekend.
On all other weekdays it selects **Last 24 Hours**.

## Windows Task Scheduler setup

1. Open **Task Scheduler** → *Create Basic Task*
2. **Name:** `Mergermarket Newsletter`
3. **Trigger:** Daily at **08:45**; repeat Mon–Fri only  
   *(Advanced settings → "Run task as soon as possible after a scheduled start is missed")*
4. **Action:** Start a program → browse to `run_mergermarket.bat`
5. **Settings:** *Run only when user is logged on*

The batch file appends all output to `C:\Temp\mergermarket_log.txt`.

## Output files

| File | Description |
|---|---|
| `C:\Temp\mergermarket_raw_YYYYMMDD.xlsx` | Raw downloaded Excel report (kept for reference) |
| `C:\Temp\mergermarket_report_YYYYMMDD.docx` | Formatted Word document |
| `C:\Temp\mergermarket_log.txt` | Running log with timestamps |

## Word document structure

```
Table of Contents          ← field; right-click → Update Field after opening
--------------------------
[Report 1 heading]         ← Aptos 12 pt Bold, Heading 1 style (picked up by TOC)
[Report 1 body text]
(Top)                      ← hyperlink to top of document
--------------------------
[Report 2 heading]
...
--------------------------
N Reports
```

## Notes

- The Outlook email is **displayed but not sent automatically** — review and
  send manually.
- The Word TOC shows a placeholder until you open the `.docx` and press
  **Ctrl+A → F9** (or right-click the TOC field → *Update Field*).
- The browser runs in **headed mode** by default so you can see what is
  happening and intervene if a CAPTCHA or unexpected dialog appears.
  Add `--headless` for unattended/scheduled runs.
- If Mergermarket changes its page layout, update the CSS selectors in
  `_set_date_range`, `_select_last_24h`, `_select_geographies`, and
  `_trigger_download` in `mergermarket_newsletter.py`.
