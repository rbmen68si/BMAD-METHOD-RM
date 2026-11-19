# /// script
# requires-python = ">=3.10"
# dependencies = ["playwright"]
# ///
"""Scarica da MyTime le esportazioni Risorse e WBS come .xls grezzi.

Avvia Chrome con un profilo dedicato, lo controlla via CDP e salva
risorse-raw.xls e wbs-raw.xls nella cartella di output.
"""

import argparse
import subprocess
import sys
import time
import urllib.request
from pathlib import Path

from playwright.sync_api import sync_playwright

CHROME_PATHS = [
    r"C:\Program Files\Google\Chrome\Application\chrome.exe",
    r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
]
MYTIME_HOST = "mytime.dedagroup.it"
MONITORING_URL = f"https://{MYTIME_HOST}/monitoring"


def cdp_ready(port):
    try:
        urllib.request.urlopen(f"http://127.0.0.1:{port}/json/version", timeout=2)
        return True
    except Exception:
        return False


def launch_chrome(port, profile):
    if cdp_ready(port):
        return
    chrome = next((p for p in CHROME_PATHS if Path(p).exists()), None)
    if not chrome:
        sys.exit("Chrome non trovato")
    subprocess.Popen(
        [chrome, f"--remote-debugging-port={port}", f"--user-data-dir={profile}",
         "--no-first-run", "--no-default-browser-check", "about:blank"],
        creationflags=0x00000008,  # DETACHED_PROCESS
    )
    for _ in range(30):
        if cdp_ready(port):
            return
        time.sleep(1)
    sys.exit(f"Chrome non risponde sulla porta {port}")


def find_page(context, fragment):
    return next((p for p in context.pages if fragment in p.url), None)


def open_monitoring(context):
    """Restituisce la scheda MyTime su /monitoring."""
    page = find_page(context, f"{MYTIME_HOST}/monitoring")
    if page:
        return page
    page = context.pages[0] if context.pages else context.new_page()
    page.goto(MONITORING_URL)
    try:
        page.wait_for_url(f"**{MYTIME_HOST}/monitoring**", timeout=60000)
        return page
    except Exception:
        pass
    # Percorso lungo (SharePoint -> IL MIO LAVORO -> APPLICAZIONI -> Mytime) descritto in SKILL.md
    sys.exit("Accesso diretto a /monitoring non riuscito: segui la navigazione manuale della skill")


def export(page, mode, dal, al, target):
    page.wait_for_selector("select[name=mode]", timeout=60000)
    page.select_option("select[name=mode]", mode)
    page.select_option("select[name=view]", "daily")
    page.fill("input[name=when]", dal)
    page.fill("input[name=whento]", al)
    with page.expect_download(timeout=300000) as dl:
        page.click("button[name=btnExport]")
    dl.value.save_as(target)
    print(f"Salvato {target} ({target.stat().st_size} byte)")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dal", required=True, help="dd/mm/yyyy")
    ap.add_argument("--al", required=True, help="dd/mm/yyyy")
    ap.add_argument("--out", required=True, help="cartella di lavoro")
    ap.add_argument("--port", type=int, default=9333)
    args = ap.parse_args()

    out = Path(args.out)
    downloads = out / "downloads"
    downloads.mkdir(parents=True, exist_ok=True)
    launch_chrome(args.port, out / "chrome-profile")

    with sync_playwright() as pw:
        browser = pw.chromium.connect_over_cdp(f"http://127.0.0.1:{args.port}")
        context = browser.contexts[0]
        page = open_monitoring(context)
        export(page, "rusr", args.dal, args.al, downloads / "risorse-raw.xls")
        export(page, "rwbs", args.dal, args.al, downloads / "wbs-raw.xls")


if __name__ == "__main__":
    main()
