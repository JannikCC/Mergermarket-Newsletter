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
import tempfile
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Optional

# ---------------------------------------------------------------------------
# Paths  (no hardcoded user or drive paths — works on any Windows machine)
# ---------------------------------------------------------------------------

# Temporary files: raw Excel download, debug screenshots/JSON
TEMP_DIR = Path(tempfile.gettempdir()) / "mergermarket"

# Persistent output: Word document, log file
OUTPUT_DIR = Path.home() / "Downloads"
LOG_FILE = OUTPUT_DIR / "mergermarket_log.txt"
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
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
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


# ── Diagnostic helpers ───────────────────────────────────────────────────────

def _dump_page_state(page, label: str) -> None:
    """
    Save a full-page screenshot and a JSON DOM dump to C:\\Temp\\mm_debug_<label>.*

    Called automatically before every key interaction so that when a selector
    fails, we have complete visibility into what was actually on the page.
    Logs a human-readable summary of all selects / buttons / inputs / iframes.
    """
    import json as _json

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    base = TEMP_DIR / f"mm_debug_{label}"

    # Screenshot (always from the top-level page object)
    try:
        page.screenshot(path=str(base.with_suffix(".png")), full_page=True)
    except Exception as exc:
        log.debug(f"Screenshot failed ({label}): {exc}")

    # DOM dump (evaluate in whichever context was passed — page or frame)
    try:
        data = page.evaluate("""() => {
            const selects = Array.from(document.querySelectorAll('select')).map(s => ({
                name: s.name, id: s.id, multiple: s.multiple,
                options: Array.from(s.options).slice(0, 30).map(o => o.text.trim())
            }));
            const buttons = Array.from(document.querySelectorAll(
                'button, input[type="button"], input[type="submit"]'
            )).map(b => ({
                tag: b.tagName.toLowerCase(),
                type: b.type || '',
                id: b.id,
                name: b.name || '',
                value: (b.value || '').slice(0, 80),
                text: (b.textContent || '').trim().slice(0, 80)
            }));
            const inputs = Array.from(document.querySelectorAll('input')).map(i => ({
                type: i.type, name: i.name || '', id: i.id,
                value: (i.value || '').slice(0, 60),
                placeholder: i.placeholder || ''
            }));
            const iframes = Array.from(document.querySelectorAll('iframe')).map(f => ({
                src: f.src, id: f.id, name: f.name
            }));
            return {
                url: location.href, title: document.title,
                selects, buttons, inputs, iframes
            };
        }""")

        with open(str(base.with_suffix(".json")), "w", encoding="utf-8") as fh:
            _json.dump(data, fh, indent=2, ensure_ascii=False)

        log.info(
            f"[DIAG {label}] {data['url'][:80]}  "
            f"selects={len(data['selects'])}  "
            f"buttons={len(data['buttons'])}  "
            f"inputs={len(data['inputs'])}  "
            f"iframes={len(data['iframes'])}"
        )
        log.info(f"  → screenshot: {base}.png")
        log.info(f"  → DOM dump  : {base}.json")

        # Print every select's options so selector issues are immediately visible
        for s in data["selects"]:
            log.info(f"  <select> name={s['name']!r} id={s['id']!r} "
                     f"multiple={s['multiple']} options={s['options'][:6]}")

        # Print every button
        for b in data["buttons"]:
            log.info(f"  <{b['tag']}> type={b['type']!r} "
                     f"value={b['value']!r} text={b['text']!r} "
                     f"id={b['id']!r} name={b['name']!r}")

        # Warn about iframes — selectors won't work across frame boundaries
        for f in data["iframes"]:
            log.warning(f"  !! iframe detected: id={f['id']!r} src={f['src']!r}")

    except Exception as exc:
        log.debug(f"DOM dump failed ({label}): {exc}")


def _find_form_context(page):
    """
    Return the Playwright context (Page or Frame) that contains the search form.

    If the form lives inside an <iframe> — common in legacy enterprise apps —
    selectors run against the main page document won't find anything.
    This function checks every frame for <select> elements and returns the
    first frame that has them.  Falls back to the main page.
    """
    # Check main document first
    try:
        n = page.evaluate("() => document.querySelectorAll('select').length")
        if n > 0:
            log.info(f"Form context: main page ({n} selects found)")
            return page
    except Exception:
        pass

    # Walk child frames
    for frame in page.frames[1:]:
        try:
            n = frame.evaluate("() => document.querySelectorAll('select').length")
            if n > 0:
                log.info(f"Form context: iframe '{frame.url}' ({n} selects found)")
                return frame
        except Exception:
            continue

    log.warning(
        "Form context: no <select> elements found anywhere on the page.\n"
        "The page may still be loading, require a login, or use a non-standard layout.\n"
        "Check the screenshot in C:\\Temp\\mm_debug_01_after_login.png"
    )
    return page  # best-effort fallback


# ── Main download orchestration ──────────────────────────────────────────────

def download_mergermarket_report(
    run_date: date,
    output_path: Path,
    *,
    headless: bool = False,
) -> Path:
    """
    Navigate Mergermarket via Playwright, configure the search, and download
    the Unformatted Report Excel file to *output_path*.

    Diagnostic screenshots + DOM dumps are saved to C:\\Temp\\mm_debug_*.{png,json}
    at every key step so selector problems are immediately diagnosable.
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
        ctx_browser = browser.new_context(accept_downloads=True)
        page = ctx_browser.new_page()

        # Navigate directly to the intelligence search page.
        # If the session is not authenticated, the server redirects automatically
        # to id.ionanalytics.com/signin — _handle_login deals with that and
        # then navigates back here.  Using domcontentloaded (not networkidle)
        # so the goto does not hang while the SSO redirect chain is in flight.
        log.info(f"Navigating to: {MERGERMARKET_URL}")
        try:
            page.goto(MERGERMARKET_URL, wait_until="domcontentloaded", timeout=30_000)
        except PWTimeout:
            _dump_page_state(page, "00_timeout")
            show_error("Mergermarket – Timeout", "Could not load the Mergermarket page within 30 s.")
            browser.close()
            raise

        _handle_login(page, mm_user, mm_pass)

        # Wait for the search form, then take a diagnostic snapshot
        try:
            page.wait_for_selector("form", timeout=15_000)
        except Exception:
            pass
        page.wait_for_timeout(1_000)  # allow any late JS rendering to settle
        _dump_page_state(page, "01_after_login")

        # Determine whether the form lives in the main document or an iframe
        ctx = _find_form_context(page)

        # ── Date range ───────────────────────────────────────────────────────
        if date_from and date_to:
            log.info(f"Setting date range: {fmt_dmy(date_from)} – {fmt_dmy(date_to)}")
            _set_date_range(ctx, date_from, date_to)
        else:
            log.info("Selecting 'Last 24 Hours' …")
            _select_last_24h(ctx)

        # ── Geography ────────────────────────────────────────────────────────
        log.info("Selecting geographies: Austria, Germany, Switzerland …")
        _dump_page_state(page, "02_before_geography")
        _select_geographies(ctx, GEOGRAPHIES)

        # ── Search (use the last Search button — after the Geography section) ─
        log.info("Submitting search …")
        _dump_page_state(page, "03_before_search")
        ctx.evaluate("window.scrollTo(0, document.body.scrollHeight)")
        ctx.wait_for_timeout(300)

        search_btns = ctx.locator("input[value='Search']")
        if search_btns.count() > 0:
            search_btns.last.scroll_into_view_if_needed()
            search_btns.last.click()
            log.info(f"Clicked Search button ({search_btns.count()} found, used last)")
        else:
            clicked = _try_click(ctx, [
                "button:text-is('Search')",
                "input[type='submit']",
            ])
            if not clicked:
                _dump_page_state(page, "03b_search_not_found")
                raise RuntimeError(
                    "Search button not found. "
                    f"See C:\\Temp\\mm_debug_03b_search_not_found.png"
                )

        page.wait_for_load_state("networkidle", timeout=30_000)

        # ── Download ─────────────────────────────────────────────────────────
        log.info("Initiating report download …")
        _dump_page_state(page, "04_results_page")
        downloaded = _trigger_download(page, ctx, output_path)
        log.info(f"Report saved to: {downloaded}")

        browser.close()

    return output_path


# ── Form-interaction helpers (accept Page or Frame as `ctx`) ─────────────────

def _handle_login(page, username: str, password: str) -> None:
    """
    Handle ION Analytics SSO login for Mergermarket.

    Detection: URL-based, not DOM-based.  The SSO provider redirects to
    id.ionanalytics.com/signin?onSuccess=… before the actual Mergermarket
    page loads.  A password-field check would miss this because the field
    is injected dynamically after the username step.

    Two-step flow:
      1. Fill username/email  →  click Continue / Next / Sign in
      2. Wait for password field to appear  →  fill  →  click Sign in
      3. Wait for redirect back to *.mergermarket.com/*
    """
    url = page.url.lower()

    # Detect login page by URL (primary signal)
    on_login_page = (
        "ionanalytics.com" in url
        or "/signin" in url
        or "/login" in url
        or "/auth/" in url
        or "mergermarket.com" not in url   # still on a redirect/SSO host
    )

    if not on_login_page:
        log.info(f"Already on Mergermarket — skipping login ({page.url[:70]})")
        return

    log.info(f"Login / SSO page detected: {page.url[:80]}")

    if not username or not password:
        show_error(
            "Mergermarket – Login Required",
            "Redirected to login page but MM_USERNAME / MM_PASSWORD are not set.\n"
            "Set them as environment variables:\n"
            "  setx MM_USERNAME \"you@company.com\"\n"
            "  setx MM_PASSWORD \"yourpassword\"",
        )
        raise RuntimeError("MM_USERNAME / MM_PASSWORD not configured")

    # ── Step 1: username / email ─────────────────────────────────────────────
    for sel in [
        "input[type='email']",
        "input[name='email']",
        "input[name='username']",
        "input[name='identifier']",
        "input[type='text']",
    ]:
        if page.query_selector(sel):
            page.fill(sel, username)
            log.info(f"Filled username/email via {sel!r}")
            break
    else:
        log.warning("Could not find a username/email field on the login page")

    # Click Continue / Next — many SSO flows show the password on the next screen
    _try_click(page, [
        "button:has-text('Continue')",
        "button:has-text('Next')",
        "input[value='Continue']",
        "input[value='Next']",
        "button[type='submit']",
        "input[type='submit']",
    ])

    # Wait up to 5 s for the password field to appear (two-step flow)
    try:
        page.wait_for_selector("input[type='password']", timeout=5_000)
        log.info("Password field appeared after Continue step")
    except Exception:
        pass  # Field may already be visible (single-step form)

    # ── Step 2: password ─────────────────────────────────────────────────────
    pw_field = page.query_selector("input[type='password']")
    if pw_field:
        page.fill("input[type='password']", password)
        log.info("Filled password field")
    else:
        log.warning("Password field still not visible — login may fail")

    # Submit: try CSS selectors first, then JS scan (case-insensitive), then Enter key.
    # The Enter-key fallback is the most reliable across all SSO implementations.
    submitted = _try_click(page, [
        "button[type='submit']",
        "input[type='submit']",
        "button:has-text('Sign in')",
        "button:has-text('Sign In')",
        "button:has-text('Log in')",
        "button:has-text('Login')",
    ])

    if not submitted:
        # JS fallback: find any submit/button whose text contains a login keyword
        submitted = page.evaluate("""() => {
            const terms = ['sign in', 'log in', 'login', 'submit', 'continue'];
            // Prefer explicit submit buttons
            const submits = Array.from(
                document.querySelectorAll('button[type="submit"], input[type="submit"]')
            );
            if (submits.length) { submits[submits.length - 1].click(); return true; }
            // Fall back to any button with matching text
            for (const el of document.querySelectorAll('button, input[type="button"]')) {
                const t = (el.textContent || el.value || '').trim().toLowerCase();
                if (terms.some(kw => t.includes(kw))) { el.click(); return true; }
            }
            return false;
        }""")

    if not submitted and pw_field:
        # Last resort: press Enter in the password field
        log.info("No submit button found — pressing Enter in password field")
        pw_field.press("Enter")

    log.info("Login submitted — waiting 3 s then navigating to intelligence page …")

    # ── Navigate immediately after submit — no URL polling ───────────────────
    # The SSO redirect can take an unpredictable amount of time.  Rather than
    # waiting for a specific URL pattern (which can block for 30 s+), we
    # simply give the auth flow 3 seconds to set its cookies, then navigate
    # directly to the target page.  If the session was not actually established,
    # the page will redirect back to the SSO and the login block will re-run.
    page.wait_for_timeout(3_000)
    log.info(f"Navigating to: {MERGERMARKET_URL}")
    page.goto(MERGERMARKET_URL, wait_until="domcontentloaded", timeout=30_000)
    log.info(f"Ready: {page.url[:80]}")


def _set_date_range(ctx, date_from: date, date_to: date) -> None:
    """
    Activate the 'Date From / Date To' radio and fill in both text inputs.

    Page structure (from live screenshot):
      ● [radio]  Last 24 Hours  [<select>]
      ○ [radio]  Date From [text input]  Date To [text input]  Clear Date

    The custom-date radio is the one whose siblings do NOT include a <select>.
    """
    ctx.evaluate("""() => {
        const radios = Array.from(document.querySelectorAll('input[type="radio"]'));
        for (const r of radios) {
            let el = r.nextElementSibling;
            let hasSelect = false;
            for (let i = 0; i < 4; i++) {
                if (!el) break;
                if (el.tagName === 'SELECT') { hasSelect = true; break; }
                el = el.nextElementSibling;
            }
            if (!hasSelect) { r.click(); return; }
        }
    }""")
    ctx.wait_for_timeout(400)

    ctx.evaluate(f"""() => {{
        const inputs = Array.from(document.querySelectorAll('input[type="text"]'));
        let fromEl = null, toEl = null;
        for (const inp of inputs) {{
            const context = (inp.parentElement?.textContent ?? '') +
                            (inp.parentElement?.parentElement?.textContent ?? '');
            if (!fromEl && context.includes('Date From')) {{ fromEl = inp; continue; }}
            if (!toEl   && context.includes('Date To'))   {{ toEl   = inp; continue; }}
        }}
        if (!fromEl && inputs[0]) fromEl = inputs[0];
        if (!toEl   && inputs[1]) toEl   = inputs[1];

        const set = (el, val) => {{
            if (!el) return;
            el.value = val;
            ['input', 'change', 'blur'].forEach(ev =>
                el.dispatchEvent(new Event(ev, {{bubbles: true}})));
        }};
        set(fromEl, '{fmt_dmy(date_from)}');
        set(toEl,   '{fmt_dmy(date_to)}');
    }}""")
    log.info(f"Date range set: {fmt_dmy(date_from)} → {fmt_dmy(date_to)}")


def _select_last_24h(ctx) -> None:
    """Click the 'Last 24 Hours' radio (the one whose siblings include a <select>)."""
    ctx.evaluate("""() => {
        const radios = Array.from(document.querySelectorAll('input[type="radio"]'));
        for (const r of radios) {
            let el = r.nextElementSibling;
            for (let i = 0; i < 4; i++) {
                if (!el) break;
                if (el.tagName === 'SELECT') { r.click(); return; }
                el = el.nextElementSibling;
            }
        }
        if (radios[0]) radios[0].click();
    }""")
    ctx.wait_for_timeout(300)
    log.info("Date mode set to 'Last 24 Hours'")


def _select_geographies(ctx, countries: list[str]) -> None:
    """
    Geography uses 4 cascading <select> elements:
      Col 1: Continent → Col 2: Sub-region → Col 3: Country

    We trigger each level by setting option.selected and dispatching 'change'.
    """
    # Step 1 – select 'Europe' in the continent column
    ok1 = ctx.evaluate("""() => {
        const sels = Array.from(document.querySelectorAll('select'));
        const s = sels.find(sel => {
            const opts = Array.from(sel.options).map(o => o.text.trim());
            return opts.includes('Europe') && opts.includes('Americas');
        });
        if (!s) return false;
        Array.from(s.options).forEach(o => { o.selected = o.text.trim() === 'Europe'; });
        s.dispatchEvent(new Event('change', {bubbles: true}));
        return true;
    }""")
    if not ok1:
        log.warning(
            "Geography: continent <select> not found. "
            "Check mm_debug_02_before_geography.json — the select with options "
            "[Africa, Americas, Asia, Europe, Middle East] was not present."
        )
        return
    log.info("Geography: continent 'Europe' selected")

    try:
        ctx.wait_for_function(
            "() => Array.from(document.querySelectorAll('select option'))"
            "       .some(o => o.text.trim() === 'Western Europe')",
            timeout=6_000,
        )
    except Exception:
        ctx.wait_for_timeout(2_000)

    # Step 2 – select 'Western Europe'
    ok2 = ctx.evaluate("""() => {
        const sels = Array.from(document.querySelectorAll('select'));
        const s = sels.find(sel =>
            Array.from(sel.options).some(o => o.text.trim() === 'Western Europe')
        );
        if (!s) return false;
        Array.from(s.options).forEach(o => { o.selected = o.text.trim() === 'Western Europe'; });
        s.dispatchEvent(new Event('change', {bubbles: true}));
        return true;
    }""")
    if not ok2:
        log.warning("Geography: 'Western Europe' not found after cascading from Europe")
        return
    log.info("Geography: sub-region 'Western Europe' selected")

    try:
        ctx.wait_for_function(
            "() => Array.from(document.querySelectorAll('select option'))"
            "       .some(o => o.text.trim() === 'Germany')",
            timeout=6_000,
        )
    except Exception:
        ctx.wait_for_timeout(2_000)

    # Step 3 – multi-select target countries
    result = ctx.evaluate("""(targets) => {
        const sels = Array.from(document.querySelectorAll('select'));
        const s = sels.find(sel =>
            targets.some(c => Array.from(sel.options).some(o => o.text.trim() === c))
        );
        if (!s) return {found: false};
        const matched = [];
        Array.from(s.options).forEach(o => {
            if (targets.includes(o.text.trim())) { o.selected = true; matched.push(o.text.trim()); }
        });
        s.dispatchEvent(new Event('change', {bubbles: true}));
        return {found: true, matched};
    }""", countries)

    if not result.get("found"):
        log.warning("Geography: country <select> not found — Austria/Germany/Switzerland absent")
    else:
        matched = result.get("matched", [])
        missing = [c for c in countries if c not in matched]
        log.info(f"Geography: selected {matched}")
        if missing:
            log.warning(f"Geography: options not found for {missing}")


def _try_click(ctx, selectors: list[str]) -> bool:
    """Try each selector in order; click the first one found. Returns True on success."""
    for sel in selectors:
        try:
            elem = ctx.query_selector(sel)
            if elem:
                elem.click()
                return True
        except Exception:
            continue
    log.warning(f"None of the click selectors matched: {selectors}")
    return False


def _js_click_by_text(ctx, *phrases: str) -> bool:
    """
    Case-insensitive JS fallback: click the first <a>, <button>, or <input>
    whose visible text / value contains any of the given phrases.
    Returns True if something was clicked.
    """
    return ctx.evaluate("""(phrases) => {
        const lower = phrases.map(p => p.toLowerCase());
        const els = Array.from(document.querySelectorAll(
            'a, button, input[type="button"], input[type="submit"]'
        ));
        for (const el of els) {
            const t = (el.textContent || el.value || '').trim().toLowerCase();
            if (lower.some(p => t.includes(p))) { el.click(); return true; }
        }
        return false;
    }""", list(phrases))


def _validate_excel_download(path: Path) -> None:
    """
    Verify the downloaded file is a valid Excel (ZIP) file by checking
    the magic bytes.  ZIP archives — and therefore all .xlsx files — start
    with the two bytes  PK  (0x50 0x4B).

    If the file is not a valid ZIP, Mergermarket probably returned an HTML
    error page instead of the report.  Log a preview and raise a clear error.
    """
    try:
        with open(path, "rb") as fh:
            header = fh.read(4)
    except OSError as exc:
        raise RuntimeError(f"Download fehlgeschlagen – Datei nicht lesbar: {exc}") from exc

    if header[:2] == b"PK":
        return  # Valid ZIP / Excel

    # Not a ZIP — capture a text preview for the log
    try:
        with open(path, encoding="utf-8", errors="replace") as fh:
            snippet = fh.read(500)
    except Exception:
        snippet = repr(header)

    log.error(f"Downloaded file is not a valid Excel/ZIP. Header bytes: {header!r}")
    log.error(f"File content preview:\n{snippet}")
    raise RuntimeError(
        "Download fehlgeschlagen – Mergermarket hat keine Excel-Datei geliefert.\n"
        f"Datei-Anfang: {snippet[:200]}"
    )


def _trigger_download(page, ctx, output_path: Path):
    """
    Mergermarket three-step download flow:

      1. Results page  →  click 'Download all'
      2. Download options page  →  click #btnUnformattedDownload
      3. Wait for 'Click here to view your report' link  →  click it
         (this final click triggers the actual file download)
    """
    ctx.wait_for_timeout(2_000)

    # ── Step 1: 'Download all' on the results page ───────────────────────────
    _dump_page_state(page, "04a_before_download_all")
    clicked = (
        _try_click(ctx, [
            "a:has-text('Download all')",
            "a:has-text('Download All')",
            "button:has-text('Download all')",
            "button:has-text('Download All')",
            "input[value='Download all']",
            "input[value='Download All']",
        ])
        or _js_click_by_text(ctx, "download all")
    )
    if not clicked:
        raise RuntimeError(
            "'Download all' element not found on the results page.\n"
            "Open C:\\Temp\\mm_debug_04a_before_download_all.png"
        )
    log.info("Clicked 'Download all'")
    page.wait_for_load_state("networkidle", timeout=15_000)

    # ── Step 2: click #btnUnformattedDownload ────────────────────────────────
    _dump_page_state(page, "04b_download_options")
    log.info("Clicking #btnUnformattedDownload …")
    clicked2 = _try_click(ctx, [
        "#btnUnformattedDownload",
        "input#btnUnformattedDownload",
        "button#btnUnformattedDownload",
    ])
    if not clicked2:
        raise RuntimeError(
            "#btnUnformattedDownload not found on the download options page.\n"
            "Check C:\\Temp\\mm_debug_04b_download_options.{png,json}"
        )
    log.info("Clicked #btnUnformattedDownload — waiting for 'Click here' link …")

    # ── Step 3: locate the 'Click here to view your report' link ─────────────
    #            Find it first, then click inside expect_download so the
    #            Playwright download handler is already registered when the
    #            browser starts the file transfer.
    log.info("Waiting for 'Click here to view your report' link …")
    link_el = None
    for sel in [
        "a:has-text('Click here to view your report')",
        "a:has-text('view your report')",
        "a:has-text('click here')",
    ]:
        try:
            link_el = ctx.wait_for_selector(sel, timeout=7_000)
            if link_el:
                log.info(f"Found download link via: {sel!r}")
                break
        except Exception:
            continue

    _dump_page_state(page, "04c_view_report_link")

    if not link_el:
        raise RuntimeError(
            "'Click here to view your report' link not found after clicking "
            "#btnUnformattedDownload.\n"
            f"Check {TEMP_DIR}\\mm_debug_04c_view_report_link.png"
        )

    log.info("Clicking download link inside expect_download context …")
    with page.expect_download(timeout=90_000) as dl_info:
        link_el.click()

    download = dl_info.value
    log.info(f"Download event captured — suggested filename: {download.suggested_filename!r}")
    download.save_as(str(output_path))
    _validate_excel_download(output_path)
    log.info(f"Download saved and validated: {output_path}")
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
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    raw_xlsx = TEMP_DIR / f"mergermarket_raw_{today_str}.xlsx"        # temp
    output_docx = OUTPUT_DIR / f"mergermarket_report_{today_str}.docx"  # Downloads

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
