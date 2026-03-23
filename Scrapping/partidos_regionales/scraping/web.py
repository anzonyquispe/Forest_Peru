"""
ONPE Web Scraper - ERM 2006, 2010, 2014
========================================
This script scrapes the ONPE historical elections page to download
municipal district-level results for 2006, 2010, and 2014.

REQUIREMENTS
------------
Install dependencies:
    pip install playwright requests pandas pymupdf playwright-stealth
    python -m playwright install chromium

HOW TO RUN
----------
1. Open Chrome with remote debugging (run this in a NEW terminal):
    Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList "--remote-debugging-port=9222 --user-data-dir=C:\temp\chrome_debug"

2. In that Chrome window, manually navigate to:
    https://www.onpe.gob.pe/elecciones/historico-elecciones/
   Wait until the page loads completely (pass Cloudflare if prompted).

3. Run this script from your terminal:
    python scraper_onpe_2006_2014.py

OUTPUT
------
CSVs saved to: Forest_Peru/Scrapping/partidos_regionales/
    - ERM2006_Municipal_Distrital.csv
    - ERM2010_Municipal_Distrital.csv
    - ERM2014_Municipal_Distrital.csv

NOTES
-----
- The script connects to your already-open Chrome (port 9222) to bypass Cloudflare.
- It navigates the ONPE page automatically: selects year → process → clicks Buscar
  → clicks Datos Abiertos → extracts links from PDF → downloads CSVs.
- UBICACION_EN_CEDULA is empty in the raw CSVs for 2010 and 2014 (ONPE did not
  populate this field for years prior to 2018).
- For 2006 only: the row order within each UBIGEO/MESA is consistent across all
  mesas and corresponds to the ballot order. Position can be derived as:
      df['UBICACION_EN_CEDULA'] = df.groupby('UBIGEO').cumcount() + 1
- For 2010 and 2014: row order is arbitrary and cannot be used to derive ballot
  position. The historical results pages (web.onpe.gob.pe) no longer exist (DNS error).
  Options: (1) request data via transparency law from ONPE, (2) check Infogob JNE.
"""

import time
import re
import io
import zipfile
import requests
import fitz  # pymupdf
from pathlib import Path
from playwright.sync_api import sync_playwright

# ── Configuration ─────────────────────────────────────────────────────────────
OUT_DIR = Path(r'C:\Users\Usuario\Documents\GitHub\Forest_Peru\Scrapping\partidos_regionales')
OUT_DIR.mkdir(exist_ok=True)

HEADERS = {'User-Agent': 'Mozilla/5.0'}
YEARS = [2006, 2010, 2014]
CDP_URL = "http://localhost:9222"  # Chrome remote debugging port


# ── Helper functions ───────────────────────────────────────────────────────────

def extract_links_from_pdf(pdf_content):
    """Download PDF from ONPE and extract datosabiertos.gob.pe links."""
    doc = fitz.open(stream=pdf_content, filetype='pdf')
    links = []
    for page in doc:
        for link in page.get_links():
            uri = link.get('uri', '')
            if uri and 'datosabiertos' in uri:
                links.append(uri)
    return list(set(links))


def get_zip_url(dataset_url):
    """Visit a datosabiertos.gob.pe dataset page and find the ZIP download URL."""
    r = requests.get(dataset_url, headers=HEADERS, timeout=60)
    matches = re.findall(r'href=["\']([^"\']+\.zip)["\']', r.text, re.I)
    for m in matches:
        if m.startswith('http'):
            return m
        return 'https://www.datosabiertos.gob.pe' + m
    return None


def download_and_extract_csv(zip_url, out_path):
    """Download a ZIP file and extract the CSV inside it."""
    print(f'  Downloading...', end=' ')
    r = requests.get(zip_url, headers=HEADERS, timeout=120)
    r.raise_for_status()
    print(f'{len(r.content)/1e6:.1f} MB')
    with zipfile.ZipFile(io.BytesIO(r.content)) as z:
        with z.open(z.namelist()[0]) as f_in:
            out_path.write_bytes(f_in.read())
    print(f'  ✓ Saved to {out_path}')


# ── Main scraping function ─────────────────────────────────────────────────────

def scrape_onpe(years):
    """
    For each year:
      1. Navigate to ONPE historico-elecciones page
      2. Select year → wait for AJAX to load processes
      3. Select ERM process → click Buscar
      4. Click Datos Abiertos → get PDF URL
      5. Extract datosabiertos.gob.pe links from PDF
    Returns dict: {year: [url1, url2, ...]}
    """
    results = {}

    with sync_playwright() as p:
        # Connect to already-open Chrome (bypasses Cloudflare)
        browser = p.chromium.connect_over_cdp(CDP_URL)
        context = browser.contexts[0]
        page = context.pages[0]

        for year in years:
            print(f'\n=== {year} ===')

            # Navigate to historical elections page
            page.goto('https://www.onpe.gob.pe/elecciones/historico-elecciones/')
            page.wait_for_load_state('domcontentloaded')
            time.sleep(3)

            # Select year and wait for AJAX to update the process dropdown
            options_before = page.inner_text('#cboElec')
            page.select_option('#cboAnio', str(year))

            print('  Waiting for AJAX...', end=' ')
            for _ in range(20):
                time.sleep(0.5)
                options_now = page.inner_text('#cboElec')
                if options_now != options_before:
                    break
            print('done')

            # Find and select the ERM (Regionales y Municipales) process
            options = page.query_selector_all('#cboElec option')
            process_val = None
            for opt in options:
                text = opt.inner_text()
                print(f'  Process: {text}')
                if 'Regionales' in text and 'Municipales' in text:
                    process_val = opt.get_attribute('value')
                    print(f'  → Selected: {text}')
                    break

            if not process_val:
                print(f'  ⚠️  ERM process not found for {year}')
                continue

            page.select_option('#cboElec', process_val)
            time.sleep(1)

            # Click Buscar button
            page.click('button:has-text("BUSCAR")')
            page.wait_for_load_state('domcontentloaded')
            time.sleep(3)

            # Click Datos Abiertos button (class: item elecciones03)
            datos_btn = page.query_selector('a.item.elecciones03')
            if not datos_btn:
                print(f'  ⚠️  Datos Abiertos button not found for {year}')
                continue

            # Get PDF URL and extract links
            href = datos_btn.get_attribute('href')
            if not href.startswith('http'):
                href = 'https://www.onpe.gob.pe' + href
            print(f'  PDF: {href}')

            r = requests.get(href, headers=HEADERS, timeout=30)
            links = extract_links_from_pdf(r.content)
            print(f'  Links found: {links}')
            results[year] = links

        browser.close()

    return results


# ── Entry point ────────────────────────────────────────────────────────────────

if __name__ == '__main__':
    # Step 1: Scrape ONPE page to get dataset links
    links_by_year = scrape_onpe(YEARS)

    # Step 2: Download distrit-level CSVs from datosabiertos.gob.pe
    for year, links in links_by_year.items():
        print(f'\n=== Downloading {year} ===')
        for link in links:
            if 'distrital' not in link.lower():
                continue  # only download district-level results
            print(f'  Dataset: {link}')
            zip_url = get_zip_url(link)
            if not zip_url:
                print('  ⚠️  ZIP download URL not found')
                continue
            out = OUT_DIR / f'ERM{year}_Municipal_Distrital.csv'
            if not out.exists():
                download_and_extract_csv(zip_url, out)
            else:
                print(f'  [SKIP] {out.name} already exists')

    print('\n✅ Done')