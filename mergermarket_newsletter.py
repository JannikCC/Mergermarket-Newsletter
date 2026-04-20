#!/usr/bin/env python3
"""
Mergermarket Daily Newsletter Automation

Downloads the Mergermarket intelligence report, converts it to a formatted
Word document, and opens a pre-composed Outlook email for manual review.

Usage:
    python mergermarket_newsletter.py                      # full pipeline
    python mergermarket_newsletter.py --dry-run FILE.xlsx  # skip download
    python mergermarket_newsletter.py --date 2024-01-08    # override date
    python mergermarket_newsletter.py --schedule           # run at 08:45 daily
"""

from __future__ import annotations

import argparse
import ctypes
import logging
import os
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

TEMP_DIR = Path(r"C:\Temp")
LOG_FILE = TEMP_DIR / "mergermarket_log.txt"
SEPARATOR = "--------------------------"
MERGERMARKET_URL = "https://www.mergermarket.com/intelligence/intelligence.asp"
GEOGRAPHIES = ["Austria", "Germany", "Switzerland"]
EMAIL_RECIPIENT = "CASE_Germany"
EMAIL_INTRO = (
    "Guten Morgen,\n\n"
    "anbei ein aktueller Auszug aus Mergermarket.\n\n"
    "Mergermarket:\n"
)

# Boilerplate text patterns to skip when parsing the Excel report
BOILERPLATE_PREFIXES = ("handelsblatt", "mergermarket", "---", "===")

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------


def setup_logging() -> logging.Logger:
    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger("mergermarket")
    logger.setLevel(logging.DEBUG)
    fmt = logging.Formatter(
        "%(asctime)s [%(levelname)s] %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    fh = logging.FileHandler(LOG_FILE, encoding="utf-8")
    fh.setFormatter(fmt)
    logger.addHandler(fh)
    ch = logging.StreamHandler()
    ch.setFormatter(fmt)
    logger.addHandler(ch)
    return logger


log = setup_logging()


def show_error(title: str, message: str) -> None:
    """Show a Windows MessageBox and log the error."""
    log.error(f"{title}: {message}")
    try:
        ctypes.windll.user32.MessageBoxW(0, message, title, 0x10)  # MB_ICONERROR
    except Exception:
        pass  # Non-Windows environment


# ---------------------------------------------------------------------------
# Date utilities
# ---------------------------------------------------------------------------


def get_run_date(override: Optional[str] = None) -> date:
    if override:
        return datetime.strptime(override, "%Y-%m-%d").date()
    return date.today()


def get_date_range(run_date: date) -> tuple[Optional[date], Optional[date]]:
    """Return (date_from, date_to) for Monday; (None, None) for Tue–Fri."""
    if run_date.weekday() == 0:  # Monday
        return run_date - timedelta(days=3), run_date  # Friday → Monday
    return None, None


def fmt_dmy(d: date) -> str:
    """Format as dd/mm/yyyy for Mergermarket date fields."""
    return d.strftime("%d/%m/%Y")


# ---------------------------------------------------------------------------
# Step 1 – Browser Automation (Playwright)
# ---------------------------------------------------------------------------


def download_mergermarket_report(
    run_date: date,
    output_path: Path,
    *,
    headless: bool = False,
) -> Path:
    """
    Navigate Mergermarket via Playwright, configure the search, and download
    the Unformatted Report Excel file to *output_path*.
    """
    try:
        from playwright.sync_api import TimeoutError as PWTimeout, sync_playwright
    except ImportError:
        show_error(
            "Mergermarket – Missing dependency",
            "Playwright is not installed.\n\nRun:\n  pip install playwright\n  playwright install chromium",
        )
        sys.exit(1)

    mm_user = os.environ.get("MM_USERNAME", "")
    mm_pass = os.environ.get("MM_PASSWORD", "")
    date_from, date_to = get_date_range(run_date)

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()

        log.info("Navigating to Mergermarket intelligence page …")
        try:
            page.goto(MERGERMARKET_URL, wait_until="networkidle", timeout=30_000)
        except PWTimeout:
            show_error("Mergermarket – Timeout", "Could not load the Mergermarket page within 30 s.")
            browser.close()
            raise

        _handle_login(page, mm_user, mm_pass)

        # Wait for the search form
        page.wait_for_selector("form", timeout=15_000)

        if date_from and date_to:
            log.info(f"Setting date range: {fmt_dmy(date_from)} – {fmt_dmy(date_to)}")
            _set_date_range(page, date_from, date_to)
        else:
            log.info("Selecting 'Last 24 Hours' …")
            _select_last_24h(page)

        log.info("Selecting geographies: Austria, Germany, Switzerland …")
        _select_geographies(page, GEOGRAPHIES)

        log.info("Submitting search …")
        _try_click(page, [
            "input[value='Search']",
            "button:has-text('Search')",
            "[type='submit']",
        ])
        page.wait_for_load_state("networkidle", timeout=30_000)

        log.info("Initiating report download …")
        downloaded = _trigger_download(page, output_path)
        log.info(f"Report saved to: {downloaded}")

        browser.close()

    return output_path


def _handle_login(page, username: str, password: str) -> None:
    """Submit login form if one is present on the page."""
    if not page.query_selector("input[type='password']"):
        log.info("No login form detected — session already authenticated.")
        return
    if not username or not password:
        show_error(
            "Mergermarket – Login Required",
            "A login page was detected but MM_USERNAME / MM_PASSWORD are not set.\n"
            "Export them as environment variables and retry.",
        )
        raise RuntimeError("MM_USERNAME / MM_PASSWORD not configured")

    log.info("Login form detected — submitting credentials …")
    for sel in ["input[name='username']", "input[type='email']", "input[name='email']"]:
        if page.query_selector(sel):
            page.fill(sel, username)
            break
    page.fill("input[type='password']", password)
    _try_click(page, ["button[type='submit']", "input[type='submit']"])
    page.wait_for_load_state("networkidle", timeout=20_000)
    log.info("Login submitted.")


def _set_date_range(page, date_from: date, date_to: date) -> None:
    """Fill the 'Date from' / 'Date to' fields."""
    for sel in ["input[name='datefrom']", "input[id*='datefrom']", "input[placeholder*='From']"]:
        if page.query_selector(sel):
            page.fill(sel, fmt_dmy(date_from))
            break
    for sel in ["input[name='dateto']", "input[id*='dateto']", "input[placeholder*='To']"]:
        if page.query_selector(sel):
            page.fill(sel, fmt_dmy(date_to))
            break


def _select_last_24h(page) -> None:
    """Choose 'Last 24 Hours' from the date-range dropdown."""
    for sel in ["select[name='daterange']", "select[id*='daterange']", "select[name*='period']"]:
        if page.query_selector(sel):
            page.select_option(sel, label="Last 24 Hours")
            return
    # Fallback: click a radio/option matching the label text
    _try_click(page, ["text=Last 24 Hours", "label:has-text('Last 24 Hours')"])


def _select_geographies(page, countries: list[str]) -> None:
    """Ctrl+click the named countries in the geography multi-select."""
    # Try to expand Western Europe section first
    for label in ["Western Europe", "Europe"]:
        expander = page.query_selector(f"text={label}")
        if expander:
            expander.click()
            time.sleep(0.4)
            break

    for idx, country in enumerate(countries):
        # Try different element types that could represent the option
        for pattern in [
            f"label:has-text('{country}')",
            f"input[value='{country}']",
            f"option:has-text('{country}')",
            f"li:has-text('{country}')",
        ]:
            locator = page.locator(pattern)
            if locator.count() > 0:
                mods = ["Control"] if idx > 0 else []
                locator.first.click(modifiers=mods)
                break
        else:
            log.warning(f"Geography option not found: {country}")


def _try_click(page, selectors: list[str]) -> bool:
    """Try each selector in order; click the first one found. Returns True on success."""
    for sel in selectors:
        try:
            elem = page.query_selector(sel)
            if elem:
                elem.click()
                return True
        except Exception:
            continue
    log.warning(f"None of the click selectors matched: {selectors}")
    return False


def _trigger_download(page, output_path: Path):
    """Click through the download dialog and capture the resulting file."""
    # Click "Download all"
    _try_click(page, [
        "text=Download all",
        "a:has-text('Download all')",
        "button:has-text('Download all')",
    ])
    page.wait_for_load_state("networkidle", timeout=15_000)

    # Select "Unformatted Report" in the Other Reports section
    _try_click(page, [
        "text=Unformatted Report",
        "label:has-text('Unformatted Report')",
        "input[value='Unformatted Report']",
    ])

    # Click the Download button
    _try_click(page, [
        "button:has-text('Download')",
        "input[value='Download']",
        "a:has-text('Download')",
    ])
    page.wait_for_load_state("networkidle", timeout=15_000)

    # "Click here to view your report" triggers the actual file download
    with page.expect_download(timeout=90_000) as dl_info:
        _try_click(page, [
            "text=Click here to view your report",
            "a:has-text('view your report')",
            "a:has-text('click here')",
        ])

    download = dl_info.value
    download.save_as(str(output_path))
    return output_path


# ---------------------------------------------------------------------------
# Step 2 – Parse the Raw Excel Report
# ---------------------------------------------------------------------------


def parse_excel_report(xlsx_path: Path) -> list[str]:
    """
    Read column E (rows 1–500) of the Mergermarket unformatted report.

    Mergermarket's unformatted Excel often begins with a header block
    (Handelsblatt date, MERGERMARKET title) and an auto-generated TOC before
    the first separator line.  Everything up to and including the first
    separator is discarded; each subsequent non-empty cell is one report entry.

    If no separator is found the entire non-boilerplate content is returned,
    so the function works even when the report layout changes.
    """
    try:
        import openpyxl
    except ImportError:
        show_error("Mergermarket – Missing dependency", "openpyxl is not installed.\nRun: pip install openpyxl")
        raise

    log.info(f"Parsing Excel report: {xlsx_path}")
    wb = openpyxl.load_workbook(str(xlsx_path), read_only=True, data_only=True)
    ws = wb.active

    all_cells: list[tuple[str, bool]] = []  # (text, after_first_separator)
    seen_separator = False

    for row_idx in range(1, 501):
        val = ws.cell(row=row_idx, column=5).value
        if val is None:
            continue
        text = str(val).strip()
        if not text:
            continue

        lower = text.lower()
        # Skip known boilerplate lines
        if any(lower.startswith(pfx) for pfx in BOILERPLATE_PREFIXES):
            if lower.startswith("---") or lower.startswith("==="):
                seen_separator = True
            continue

        all_cells.append((text, seen_separator))

    wb.close()

    # Prefer entries that came after the first separator (skip the TOC block).
    post_sep = [t for t, after in all_cells if after]
    entries = post_sep if post_sep else [t for t, _ in all_cells]

    log.info(f"Found {len(entries)} report entries in column E.")
    return entries


# ---------------------------------------------------------------------------
# Step 3 – Generate the Formatted Word Document
# ---------------------------------------------------------------------------


def _add_hyperlink_to_top(paragraph) -> None:
    """Append a '(Top)' hyperlink pointing to the _top bookmark."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), "_top")
    hyperlink.set(qn("w:history"), "1")

    r = OxmlElement("w:r")
    rPr = OxmlElement("w:rPr")
    rStyle = OxmlElement("w:rStyle")
    rStyle.set(qn("w:val"), "Hyperlink")
    rPr.append(rStyle)
    r.append(rPr)

    t = OxmlElement("w:t")
    t.text = "(Top)"
    r.append(t)
    hyperlink.append(r)
    paragraph._p.append(hyperlink)


def _insert_toc(document) -> None:
    """
    Insert a TOC field (Heading 1 only, hyperlinked, no page numbers) at the
    very beginning of the document body.
    """
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    body = document.element.body

    # TOC label paragraph
    label_p = OxmlElement("w:p")
    label_r = OxmlElement("w:r")
    label_rPr = OxmlElement("w:rPr")
    label_b = OxmlElement("w:b")
    label_rPr.append(label_b)
    label_r.append(label_rPr)
    label_t = OxmlElement("w:t")
    label_t.text = "Table of Contents"
    label_r.append(label_t)
    label_p.append(label_r)

    # TOC field paragraph
    # Switches: \h = hyperlinks, \z = hide in web layout, \n = no page numbers,
    #           \o "1-1" = include only Heading 1
    field_p = OxmlElement("w:p")
    field_r = OxmlElement("w:r")

    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    field_r.append(begin)

    instr_r = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = ' TOC \\h \\z \\n \\o "1-1" '
    instr_r.append(instr)

    sep_r = OxmlElement("w:r")
    sep_char = OxmlElement("w:fldChar")
    sep_char.set(qn("w:fldCharType"), "separate")
    sep_r.append(sep_char)
    placeholder_r = OxmlElement("w:r")
    placeholder_t = OxmlElement("w:t")
    placeholder_t.text = "[Right-click → Update Field to refresh this Table of Contents]"
    placeholder_r.append(placeholder_t)

    end_r = OxmlElement("w:r")
    end_char = OxmlElement("w:fldChar")
    end_char.set(qn("w:fldCharType"), "end")
    end_r.append(end_char)

    field_p.append(field_r)
    field_p.append(instr_r)
    field_p.append(sep_r)
    field_p.append(placeholder_r)
    field_p.append(end_r)

    # Insert both paragraphs before all existing body content
    body.insert(0, field_p)
    body.insert(0, label_p)


def _apply_heading_formatting(paragraph, font_name: str = "Aptos") -> None:
    """Set Heading 1 style then override with Aptos 12pt Bold Black."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Pt, RGBColor

    paragraph.style = "Heading 1"
    for run in paragraph.runs:
        run.font.name = font_name
        run.font.size = Pt(12)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
        # Ensure the font name is stored in the rFonts XML element
        rPr = run._r.get_or_add_rPr()
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        rFonts.set(qn("w:ascii"), font_name)
        rFonts.set(qn("w:hAnsi"), font_name)


def generate_word_document(
    entries: list[str],
    output_path: Path,
    run_date: date,
) -> Path:
    """
    Build the formatted .docx from a list of report entry strings.

    Each string in *entries* represents one intelligence report; its first
    non-empty line becomes the Heading 1 and subsequent lines become body text.
    """
    try:
        from docx import Document
    except ImportError:
        show_error("Mergermarket – Missing dependency", "python-docx is not installed.\nRun: pip install python-docx")
        raise

    log.info(f"Generating Word document: {output_path}")

    # Prefer Aptos (Office 2024+); fall back to Calibri
    heading_font = "Aptos"

    doc = Document()

    # Remove the single empty paragraph that python-docx adds by default
    for p in list(doc.paragraphs):
        p._element.getparent().remove(p._element)

    report_count = 0

    for entry in entries:
        lines = [ln.strip() for ln in entry.splitlines() if ln.strip()]
        if not lines:
            continue

        report_count += 1

        # Separator line before each report
        doc.add_paragraph(SEPARATOR).style = "Normal"

        # First line → Heading 1
        heading_para = doc.add_paragraph()
        heading_para.add_run(lines[0])
        _apply_heading_formatting(heading_para, font_name=heading_font)

        # Remaining lines → Normal body text
        for line in lines[1:]:
            # Skip any "(Top)" the raw data might already contain
            if line.strip() == "(Top)":
                continue
            doc.add_paragraph(line).style = "Normal"

        # (Top) hyperlink at the end of each report
        top_para = doc.add_paragraph()
        top_para.style = "Normal"
        _add_hyperlink_to_top(top_para)

    if report_count == 0:
        log.warning("No report entries — the Word document will be empty.")
        show_error(
            "Mergermarket – No Reports",
            "No intelligence reports were found for today.\nCheck the downloaded Excel file.",
        )

    # Trailing separator and report count
    doc.add_paragraph(SEPARATOR).style = "Normal"
    doc.add_paragraph(f"{report_count} Reports").style = "Normal"

    # Table of Contents at the very top
    _insert_toc(doc)

    doc.save(str(output_path))
    log.info(f"Word document saved: {output_path}  ({report_count} reports)")
    return output_path


# ---------------------------------------------------------------------------
# Step 4 – Compose and Display the Outlook Email
# ---------------------------------------------------------------------------


def compose_outlook_email(word_doc_path: Path, run_date: date) -> None:
    """
    Create a new Outlook MailItem, insert the intro text and the content of
    the Word document, then display it so the user can review before sending.
    """
    try:
        import win32com.client
    except ImportError:
        show_error(
            "Mergermarket – Missing dependency",
            "pywin32 is not installed.\nRun: pip install pywin32",
        )
        raise

    subject = f"Mergermarket {run_date.strftime('%d.%m.%Y')}"
    doc_path_str = str(word_doc_path.resolve())

    log.info("Launching Outlook …")
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        show_error("Mergermarket – Outlook Error", f"Could not connect to Outlook:\n{exc}")
        raise

    log.info("Opening Word document via COM …")
    try:
        word_app = win32com.client.Dispatch("Word.Application")
        word_app.Visible = False
        source_doc = word_app.Documents.Open(doc_path_str, ReadOnly=True)
    except Exception as exc:
        show_error("Mergermarket – Word Error", f"Could not open the Word document:\n{exc}")
        raise

    try:
        mail = outlook.CreateItem(0)  # 0 = olMailItem
        mail.Subject = subject
        mail.To = EMAIL_RECIPIENT

        # Display the email first — WordEditor is only available after Display()
        mail.Display()

        # Get the embedded Word editor for the message body
        inspector = mail.GetInspector
        mail_doc = inspector.WordEditor          # Word.Document COM object
        word_selection = mail_doc.Application.Selection

        # Move to the very start of the body and insert the intro block
        word_selection.HomeKey(Unit=6)  # wdStory = 6
        for line in EMAIL_INTRO.splitlines():
            word_selection.TypeText(line)
            word_selection.TypeParagraph()

        # Copy all content from the source Word document and paste it in
        source_doc.Content.Copy()
        word_selection.EndKey(Unit=6)  # move to end before pasting
        word_selection.Paste()

        log.info("Outlook email composed and displayed — ready for manual review.")

    finally:
        source_doc.Close(SaveChanges=False)
        if word_app.Documents.Count == 0:
            word_app.Quit()


# ---------------------------------------------------------------------------
# Main pipeline
# ---------------------------------------------------------------------------


def run(run_date: date, dry_run_xlsx: Optional[Path] = None, headless: bool = False) -> None:
    today_str = run_date.strftime("%Y%m%d")
    TEMP_DIR.mkdir(parents=True, exist_ok=True)

    raw_xlsx = TEMP_DIR / f"mergermarket_raw_{today_str}.xlsx"
    output_docx = TEMP_DIR / f"mergermarket_report_{today_str}.docx"

    try:
        # ── Step 1: Download ─────────────────────────────────────────────────
        if dry_run_xlsx:
            raw_xlsx = dry_run_xlsx
            log.info(f"[DRY-RUN] Skipping browser download; using: {raw_xlsx}")
        else:
            download_mergermarket_report(run_date, raw_xlsx, headless=headless)

        # ── Step 2: Parse ────────────────────────────────────────────────────
        entries = parse_excel_report(raw_xlsx)
        if not entries:
            show_error(
                "Mergermarket – No Data",
                f"No report entries were found in:\n{raw_xlsx}",
            )
            return

        # ── Step 3: Format Word document ─────────────────────────────────────
        generate_word_document(entries, output_docx, run_date)

        # ── Step 4: Compose email ────────────────────────────────────────────
        compose_outlook_email(output_docx, run_date)

    except Exception as exc:
        log.exception("Fatal error in Mergermarket newsletter automation")
        show_error(
            "Mergermarket – Fatal Error",
            f"The automation failed:\n{exc}\n\nFull log: {LOG_FILE}",
        )
        sys.exit(1)


# ---------------------------------------------------------------------------
# Scheduler  (optional — use --schedule flag or Windows Task Scheduler)
# ---------------------------------------------------------------------------


def start_scheduler(date_override: Optional[str] = None, headless: bool = False) -> None:
    """Block forever, executing the pipeline on weekdays at 08:45."""
    try:
        import schedule
    except ImportError:
        show_error("Mergermarket – Missing dependency", "schedule is not installed.\nRun: pip install schedule")
        sys.exit(1)

    def _job() -> None:
        today = date.today()
        if today.weekday() >= 5:
            log.info("Weekend — skipping scheduled run.")
            return
        log.info("=== Scheduled run triggered ===")
        run(get_run_date(date_override), headless=headless)

    schedule.every().day.at("08:45").do(_job)
    log.info("Scheduler active — waiting for 08:45 on weekdays …  (Ctrl+C to stop)")
    while True:
        schedule.run_pending()
        time.sleep(30)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Automate the Mergermarket daily newsletter workflow.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument(
        "--dry-run",
        metavar="XLSX_PATH",
        help="Skip browser download and use an existing Excel file instead.",
    )
    parser.add_argument(
        "--date",
        metavar="YYYY-MM-DD",
        help="Override today's date (e.g. for a Monday run covering the weekend).",
    )
    parser.add_argument(
        "--schedule",
        action="store_true",
        help="Run as a background scheduler (Mon–Fri at 08:45).",
    )
    parser.add_argument(
        "--headless",
        action="store_true",
        help="Run the browser in headless mode (no visible window).",
    )
    args = parser.parse_args()

    run_date = get_run_date(args.date)
    log.info(f"=== Mergermarket Newsletter  {run_date}  ({'Monday' if run_date.weekday() == 0 else run_date.strftime('%A')}) ===")

    if args.schedule:
        start_scheduler(date_override=args.date, headless=args.headless)
    elif args.dry_run:
        run(run_date, dry_run_xlsx=Path(args.dry_run), headless=args.headless)
    else:
        run(run_date, headless=args.headless)


if __name__ == "__main__":
    main()
