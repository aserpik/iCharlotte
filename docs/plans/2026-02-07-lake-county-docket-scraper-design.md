# Lake County Docket Scraper Design

**Date:** 2026-02-07
**Status:** Approved

## Overview

Build a new docket scraper for Lake County Superior Court (`portal.lake.courts.ca.gov`). The court uses the Journal Technologies eCourt Public Portal (same vendor as Riverside), but with a simpler tabbed UI and no login requirement.

## Court Portal Details

- **URL:** `https://portal.lake.courts.ca.gov/public-portal/?q=node/393`
- **Platform:** Journal Technologies eCourt on Drupal
- **Auth:** None (public access)
- **CAPTCHA:** reCAPTCHA v2 on search form
- **Layout:** Tabbed — Case, Filings, Parties, Documents, Events, Case Transfer
- **Pagination:** None observed within tabs
- **Case detail URL pattern:** `?q=node/394/{case_id}`

## Architecture

### Flow

1. Launch Playwright browser, navigate to search page
2. Enter case number in search form
3. Solve reCAPTCHA v2 via 2Captcha API
4. Submit search, click into case result
5. Iterate through all 6 tabs, rendering each to a temporary PDF
6. Merge tab PDFs into single output file
7. Clean up temp files

### reCAPTCHA Solving

- Extract `data-sitekey` from reCAPTCHA div
- Submit sitekey + page URL to 2Captcha API, poll for token
- Inject token into `g-recaptcha-response` textarea via JavaScript
- Submit the search form
- Retry up to 3 times on failure

### Stealth Measures

- Chromium args: `--disable-blink-features=AutomationControlled`, `--no-sandbox`
- User-Agent spoofing (Chrome on Windows)
- Patch `navigator.webdriver` to false

### Tab Capture & PDF Generation

- `page.emulate_media(media="screen")` to prevent print CSS issues
- For each tab: click tab, wait for `networkidle` + short delay
- Inject CSS: `.collapse { display: block !important }` to force-expand collapsible sections
- Render via `page.pdf(format="Letter", print_background=True)`
- Merge all tab PDFs with `pypdf.PdfWriter`
- Output: `docket_{case_number}_{YYYY.MM.DD}.pdf`

### Error Handling

- reCAPTCHA failure: retry 3x, then exit code 1
- Case not found: print message, exit code 1
- Tab load timeout: skip tab, continue with remaining
- Browser crash: caught by orchestrator retry logic

## Integration

### Orchestrator (`docket.py`)

- Add `"Lake"` to county routing map pointing to `lake_docket_scraper.py`
- No extra arguments needed

### CLI Interface

```
python lake_docket_scraper.py <case_number> [--headful]
```

### Dependencies

- `playwright` (async) — browser automation
- `urllib` — 2Captcha HTTP API
- `pypdf.PdfWriter` — PDF merging
- `dotenv` — for `TWOCAPTCHA_API_KEY`

## File

`Scripts/lake_docket_scraper.py`
