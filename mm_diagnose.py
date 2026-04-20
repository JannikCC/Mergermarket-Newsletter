"""
Mergermarket Selector Diagnostic Tool
======================================
Run this ONCE on your Windows machine to capture the exact HTML structure
of the Mergermarket search page and download dialogs.

Output saved to: C:\\Temp\\mm_diagnose\\
  - 01_search_page.png
  - 02_after_search.png
  - 03_download_dialog.png
  - elements.json     ← share this with Claude to fix the selectors

Usage:
    python mm_diagnose.py
    python mm_diagnose.py --user you@example.com --password secret
"""

import argparse
import json
import os
import time
from datetime import date
from pathlib import Path

OUT_DIR = Path(r"C:\Temp\mm_diagnose")
MM_URL = "https://www.mergermarket.com/intelligence/intelligence.asp"


def dump_elements(page, stage: str) -> dict:
    """Capture all interactive elements on the current page via JavaScript."""
    data = page.evaluate("""() => {
        function getLabel(el) {
            if (el.labels && el.labels[0]) return el.labels[0].textContent.trim();
            const lbl = el.closest('label');
            if (lbl) return lbl.textContent.trim().replace(el.value || '', '').trim();
            const sib = el.previousElementSibling;
            if (sib) return sib.textContent.trim();
            return '';
        }

        const result = { inputs: [], selects: [], buttons: [], links: [], checkboxes: [] };

        document.querySelectorAll('input').forEach(el => {
            result.inputs.push({
                type: el.type, name: el.name, id: el.id,
                value: el.value, placeholder: el.placeholder,
                label: getLabel(el),
                html: el.outerHTML.slice(0, 400)
            });
        });

        document.querySelectorAll('select').forEach(el => {
            result.selects.push({
                name: el.name, id: el.id, multiple: el.multiple,
                options: Array.from(el.options).map(o => ({v: o.value, t: o.text.trim()})),
                html: el.outerHTML.slice(0, 600)
            });
        });

        document.querySelectorAll('button, input[type="button"], input[type="submit"]').forEach(el => {
            result.buttons.push({
                tag: el.tagName, name: el.name, id: el.id,
                text: (el.textContent || el.value || '').trim(),
                html: el.outerHTML.slice(0, 300)
            });
        });

        document.querySelectorAll('a').forEach(el => {
            const text = el.textContent.trim();
            if (text && text.length < 120) {
                result.links.push({
                    text: text, href: el.href, id: el.id,
                    html: el.outerHTML.slice(0, 300)
                });
            }
        });

        document.querySelectorAll('input[type="checkbox"], input[type="radio"]').forEach(el => {
            result.checkboxes.push({
                type: el.type, name: el.name, id: el.id, value: el.value,
                checked: el.checked, label: getLabel(el),
                html: el.outerHTML.slice(0, 400)
            });
        });

        // Capture any div/span/li that looks like a clickable geography item
        const geoKeywords = ['Austria', 'Germany', 'Switzerland', 'Western Europe', 'Europe'];
        result.geo_candidates = [];
        geoKeywords.forEach(kw => {
            document.querySelectorAll('*').forEach(el => {
                if (el.children.length === 0 && el.textContent.trim() === kw) {
                    result.geo_candidates.push({
                        keyword: kw,
                        tag: el.tagName,
                        id: el.id,
                        className: el.className,
                        name: el.name || '',
                        parent_tag: el.parentElement ? el.parentElement.tagName : '',
                        parent_id: el.parentElement ? el.parentElement.id : '',
                        parent_class: el.parentElement ? el.parentElement.className : '',
                        html: el.outerHTML.slice(0, 400)
                    });
                }
            });
        });

        return result;
    }""")
    return {"stage": stage, "url": page.url, "title": page.title(), "elements": data}


def run_diagnostic(username: str, password: str) -> None:
    from playwright.sync_api import sync_playwright

    OUT_DIR.mkdir(parents=True, exist_ok=True)
    results = []

    with sync_playwright() as pw:
        browser = pw.chromium.launch(headless=False)  # headful so you can see what's happening
        page = browser.new_page()

        # ── Navigate ────────────────────────────────────────────────────────
        print(f"[1/6] Navigating to {MM_URL} …")
        page.goto(MM_URL, wait_until="domcontentloaded", timeout=30_000)
        time.sleep(2)

        # ── Login if needed ─────────────────────────────────────────────────
        if page.query_selector("input[type='password']"):
            print("[2/6] Login page detected — entering credentials …")
            if not username or not password:
                print("      !! No credentials provided. Pass --user and --password, or log in manually.")
                print("      Waiting 30 s for manual login …")
                time.sleep(30)
            else:
                for sel in ["input[name='username']", "input[type='email']", "input[name='email']"]:
                    if page.query_selector(sel):
                        page.fill(sel, username)
                        break
                page.fill("input[type='password']", password)
                for sel in ["button[type='submit']", "input[type='submit']"]:
                    if page.query_selector(sel):
                        page.click(sel)
                        break
                page.wait_for_load_state("networkidle", timeout=20_000)
                print("      Login submitted.")
        else:
            print("[2/6] No login form — already authenticated.")

        # ── Search page ──────────────────────────────────────────────────────
        print("[3/6] Capturing search page …")
        time.sleep(2)
        page.screenshot(path=str(OUT_DIR / "01_search_page.png"), full_page=True)
        results.append(dump_elements(page, "search_page"))
        print(f"      Screenshot: {OUT_DIR / '01_search_page.png'}")

        # ── Click Search (try common patterns) ───────────────────────────────
        print("[4/6] Attempting to click Search …")
        clicked = False
        for sel in [
            "input[value='Search']",
            "button:has-text('Search')",
            "input[type='submit'][value*='Search']",
            "[type='submit']",
        ]:
            try:
                elem = page.query_selector(sel)
                if elem:
                    elem.click()
                    page.wait_for_load_state("networkidle", timeout=20_000)
                    print(f"      Clicked via: {sel}")
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            print("      Could not click Search automatically — please click it manually.")
            print("      Waiting 15 s …")
            time.sleep(15)

        time.sleep(2)
        page.screenshot(path=str(OUT_DIR / "02_after_search.png"), full_page=True)
        results.append(dump_elements(page, "after_search"))
        print(f"      Screenshot: {OUT_DIR / '02_after_search.png'}")

        # ── Click Download all ───────────────────────────────────────────────
        print("[5/6] Attempting to click 'Download all' …")
        clicked = False
        for sel in [
            "text=Download all",
            "a:has-text('Download all')",
            "button:has-text('Download all')",
            "input[value*='Download']",
        ]:
            try:
                elem = page.query_selector(sel)
                if elem:
                    elem.click()
                    time.sleep(3)
                    print(f"      Clicked via: {sel}")
                    clicked = True
                    break
            except Exception:
                continue

        if not clicked:
            print("      Could not click 'Download all' — please click it manually.")
            print("      Waiting 15 s …")
            time.sleep(15)

        page.screenshot(path=str(OUT_DIR / "03_download_dialog.png"), full_page=True)
        results.append(dump_elements(page, "download_dialog"))
        print(f"      Screenshot: {OUT_DIR / '03_download_dialog.png'}")

        # ── Save JSON ────────────────────────────────────────────────────────
        print("[6/6] Saving element data …")
        out_json = OUT_DIR / "elements.json"
        with open(out_json, "w", encoding="utf-8") as f:
            json.dump(results, f, indent=2, ensure_ascii=False)
        print(f"\n✓ Diagnostic complete.\n")
        print(f"  Screenshots  → {OUT_DIR}\\*.png")
        print(f"  Element data → {out_json}")
        print(f"\nPlease share the contents of {out_json} and the screenshots with Claude.")

        input("\nPress Enter to close the browser …")
        browser.close()


def main():
    parser = argparse.ArgumentParser(description="Mergermarket selector diagnostic tool")
    parser.add_argument("--user", default=os.environ.get("MM_USERNAME", ""),
                        help="Mergermarket username (or set MM_USERNAME env var)")
    parser.add_argument("--password", default=os.environ.get("MM_PASSWORD", ""),
                        help="Mergermarket password (or set MM_PASSWORD env var)")
    args = parser.parse_args()
    run_diagnostic(args.user, args.password)


if __name__ == "__main__":
    main()
