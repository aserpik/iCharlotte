import sys
import asyncio
import time
import re
import os
from google import genai
from playwright.async_api import async_playwright

async def solve_image_captcha(page):
    """Automatically extracts, solves, and fills the captcha using Gemini Vision."""
    debug_log = open("riverside_debug.log", "a")
    def log_debug(msg):
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        debug_log.write(f"[{timestamp}] {msg}\n")
        debug_log.flush()
        print(msg)

    try:
        log_debug("Attempting to solve CAPTCHA with Gemini Vision...")
        
        # Strategy 1: Look for specific Image CAPTCHA (img tag)
        captcha_img = page.locator("img[alt='Image CAPTCHA']").first
        is_image_captcha = False
        
        if await captcha_img.count() > 0:
            log_debug("Found Image CAPTCHA.")
            target_element = captcha_img
            is_image_captcha = True
        else:
            log_debug("Image CAPTCHA not found. Checking for Math/Text CAPTCHA...")
            # Strategy 2: Look for Math/Text container
            target_element = page.locator(".form-item-captcha-response").first
            
            if await target_element.count() == 0:
                log_debug("Warning: Could not find any captcha container.")
                return False
            
            # Check if it looks like a math problem
            text = await target_element.inner_text()
            log_debug(f"Math/Text Captcha Content: {text}")

        # Screenshot the target
        image_path = "captcha_challenge.png"
        await target_element.screenshot(path=image_path)
        log_debug(f"Captured captcha to {image_path}")

        # Configure Gemini
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            log_debug("Error: GEMINI_API_KEY not found in environment variables.")
            return False
            
        client = genai.Client(api_key=api_key)
        # Switch to gemini-2.0-flash for potentially better OCR on distorted images
        
        sample_file = None
        try:
            with open(image_path, "rb") as f:
                sample_file = client.files.upload(file=f, config={'display_name': "Captcha Image", 'mime_type': 'image/png'})
        finally:
            # Delete image after upload attempt
            if os.path.exists(image_path):
                os.remove(image_path)
        
        if not sample_file:
            return False

        # Tailored prompt
        if is_image_captcha:
            prompt = """Look at this CAPTCHA image and return ONLY the exact alphanumeric characters shown.

Rules:
- Output ONLY the characters, nothing else (no quotes, no explanation)
- Be case-sensitive (distinguish uppercase A from lowercase a)
- Common confusions to watch for:
  * 0 (zero) vs O (letter O) vs o (lowercase o)
  * 1 (one) vs l (lowercase L) vs I (uppercase i)
  * 5 vs S vs s
  * 8 vs B vs g
  * 2 vs Z vs z
  * 6 vs b vs G
  * 9 vs q vs g
- Look at the shape and style carefully
- Most CAPTCHAs are 4-6 characters long

Output the characters now:"""
        else:
            prompt = "Solve this math problem (e.g., '1 + 5 ='). Return ONLY the result number."
        
        # Use gemini-3-pro-preview for best OCR accuracy on distorted text
        response = client.models.generate_content(model="gemini-3-pro-preview", contents=[sample_file, prompt])
        captcha_text = response.text.strip()

        # Handle verbose responses - extract just the alphanumeric characters
        # If response contains newlines or is too long, it's probably an explanation
        if '\n' in captcha_text or len(captcha_text) > 10:
            # Try to find a short alphanumeric sequence (typical CAPTCHA is 4-6 chars)
            import re as regex
            # Look for standalone alphanumeric sequences of 4-7 characters
            matches = regex.findall(r'\b[A-Za-z0-9]{4,7}\b', captcha_text)
            if matches:
                # Take the last match (often the final answer)
                captcha_text = matches[-1]
                log_debug(f"Extracted CAPTCHA from verbose response: '{captcha_text}'")

        # Remove any quotes or extra whitespace that the model might add
        captcha_text = captcha_text.strip("'\"` \n\r")
        
        log_debug(f"Gemini solved captcha: '{captcha_text}'")
        
        if captcha_text:
            await page.fill("#edit-captcha-response", captcha_text)
            await page.click("#edit-submit")
            return True
            
    except Exception as e:
        log_debug(f"Error solving captcha with Gemini: {e}")
    
    return False

async def solve_math_captcha(page):
    """Automatically extracts, solves, and fills the math captcha."""
    try:
        # The math problem is text inside the parent div, NOT inside the label
        # structure: <div class="... form-item-captcha-response ..."> <label>...</label> 1 + 9 = <input> ... </div>
        captcha_container = page.locator(".form-item-captcha-response")
        container_text = await captcha_container.inner_text()
        
        print(f"Debug: Captcha container text: '{container_text}'")
        
        # Look for pattern like "1 + 9 =" or "1 + 9"
        # We look for two numbers separated by a plus sign
        match = re.search(r'(\d+)\s*\+\s*(\d+)', container_text)
        
        if match:
            num1 = int(match.group(1))
            num2 = int(match.group(2))
            result = num1 + num2
            await page.fill("#edit-captcha-response", str(result))
            print(f"Solved math captcha: {num1} + {num2} = {result}")
            return True
        else:
            print("Warning: Could not find math pattern in captcha text.")
            
    except Exception as e:
        print(f"Warning: Could not solve math captcha automatically: {e}")
    return False

async def expand_all_pagination(page):
    """
    Expands all paginated sections by collecting ALL data from ALL pages,
    then rebuilding ALL tables at once right before returning.

    Two-phase approach prevents AJAX from one section overwriting another:
      Phase 1 (COLLECT): Navigate pages and collect row HTML for each section
      Phase 2 (REBUILD): Replace all tables with full data (no more AJAX after this)
    """

    # JS helpers ----------------------------------------------------------
    FIND_TABLE_ROWS_JS = """(expectedHeaders) => {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            let ths = Array.from(table.querySelectorAll('thead th'));
            if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
            const hdr = ths.map(th => th.textContent.trim());
            if (expectedHeaders.length > 0 &&
                expectedHeaders.every(h => hdr.some(th => th === h))) {
                const tbody = table.querySelector('tbody');
                if (!tbody) continue;
                const rows = Array.from(tbody.querySelectorAll(':scope > tr'))
                    .filter(tr => !tr.querySelector('nav') &&
                                  !/Results\\s*\\d/.test(tr.textContent));
                return { rows: rows.map(tr => tr.outerHTML), headers: hdr };
            }
        }
        return null;
    }"""

    FINGERPRINT_JS = """(expectedHeaders) => {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            let ths = Array.from(table.querySelectorAll('thead th'));
            if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
            const hdr = ths.map(th => th.textContent.trim());
            if (expectedHeaders.length > 0 &&
                expectedHeaders.every(h => hdr.some(th => th === h))) {
                const tbody = table.querySelector('tbody');
                if (!tbody) continue;
                const first = Array.from(tbody.querySelectorAll(':scope > tr'))
                    .find(tr => !tr.querySelector('nav') &&
                                !/Results\\s*\\d/.test(tr.textContent));
                return first ? first.textContent.trim().substring(0, 200) : '';
            }
        }
        return '';
    }"""

    WAIT_CHANGE_JS = """([expectedHeaders, oldFp]) => {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            let ths = Array.from(table.querySelectorAll('thead th'));
            if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
            const hdr = ths.map(th => th.textContent.trim());
            if (expectedHeaders.length > 0 &&
                expectedHeaders.every(h => hdr.some(th => th === h))) {
                const tbody = table.querySelector('tbody');
                if (!tbody) return true;
                const first = Array.from(tbody.querySelectorAll(':scope > tr'))
                    .find(tr => !tr.querySelector('nav') &&
                                !/Results\\s*\\d/.test(tr.textContent));
                return (first ? first.textContent.trim().substring(0, 200) : '') !== oldFp;
            }
        }
        return false;
    }"""

    REDISCOVER_FUNC_JS = """(expectedHeaders) => {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            let ths = Array.from(table.querySelectorAll('thead th'));
            if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
            const hdr = ths.map(th => th.textContent.trim());
            if (expectedHeaders.length > 0 &&
                expectedHeaders.every(h => hdr.some(th => th === h))) {
                for (const link of table.querySelectorAll('a[onclick*="_pageLoad"]')) {
                    const m = (link.getAttribute('onclick') || '')
                        .match(/(\\w+_pageLoad)\\(\\s*'([^']*)'/);
                    if (m) return { funcName: m[1], viewId: m[2] };
                }
            }
        }
        return null;
    }"""

    # Full rediscovery: returns funcName, viewId, pageSize, AND maxPage for a section
    REDISCOVER_FULL_JS = """(expectedHeaders) => {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            let ths = Array.from(table.querySelectorAll('thead th'));
            if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
            const hdr = ths.map(th => th.textContent.trim());
            if (expectedHeaders.length > 0 &&
                expectedHeaders.every(h => hdr.some(th => th === h))) {
                let maxPage = 0;
                let funcName = null;
                let viewId = null;
                let pageSize = null;
                for (const link of table.querySelectorAll('a[onclick*="_pageLoad"]')) {
                    const onclick = link.getAttribute('onclick') || '';
                    const m = onclick.match(/(\\w+_pageLoad)\\(\\s*'([^']*)'\\s*,\\s*(\\d+)\\s*,\\s*(\\d+)/);
                    if (m) {
                        funcName = m[1];
                        viewId = m[2];
                        pageSize = parseInt(m[4]);
                    }
                    const pn = parseInt(link.textContent.trim());
                    if (!isNaN(pn) && pn > maxPage) maxPage = pn;
                }
                if (funcName) {
                    return { funcName, viewId, pageSize, maxPage };
                }
            }
        }
        return null;
    }"""

    REEXPAND_SECTIONS_JS = """() => {
        document.querySelectorAll('.collapse, .panel-collapse, [class*="collapse"]').forEach(el => {
            el.style.display = 'block';
            el.style.height = 'auto';
            el.style.overflow = 'visible';
            el.classList.add('in', 'show');
        });
    }"""

    try:
        # ============================================================
        # PHASE 1: DISCOVER & COLLECT rows from all paginated sections
        # ============================================================
        groups = await page.evaluate("""() => {
            const links = document.querySelectorAll('a[onclick*="_pageLoad"]');
            const gm = {};
            for (const link of links) {
                const onclick = link.getAttribute('onclick') || '';
                const fm = onclick.match(/^(\\w+_pageLoad)\\(/);
                if (!fm) continue;
                const fn = fm[1];
                const pn = parseInt(link.textContent.trim());
                if (isNaN(pn)) continue;
                if (!gm[fn]) gm[fn] = { funcName: fn, maxPage: 1, headers: [], sampleOnclick: '' };
                if (pn > gm[fn].maxPage) gm[fn].maxPage = pn;
                if (pn >= 2 || !gm[fn].sampleOnclick) gm[fn].sampleOnclick = onclick;
            }
            for (const key of Object.keys(gm)) {
                const g = gm[key];
                const fl = Array.from(document.querySelectorAll('a[onclick*="_pageLoad"]'))
                    .filter(a => (a.getAttribute('onclick') || '').includes(g.funcName));
                if (fl.length === 0) continue;
                const table = fl[0].closest('table');
                if (table) {
                    let ths = Array.from(table.querySelectorAll('thead th'));
                    if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
                    g.headers = ths.map(th => th.textContent.trim());
                }
            }
            return Object.values(gm);
        }""")

        if not groups:
            print("No pagination found on this page.")
            return

        print(f"Found {len(groups)} paginated section(s)")

        # collected_data: list of { headers, rows_html } for rebuild phase
        collected_data = []

        for group in groups:
            headers = group.get('headers', [])

            if not headers:
                continue

            # Re-expand all accordion sections before each group.
            # AJAX from processing the previous section may have collapsed this one.
            await page.evaluate(REEXPAND_SECTIONS_JS)
            await page.wait_for_timeout(500)

            # Re-discover pagination info FRESH from the current DOM.
            # The initial discovery funcNames are stale after any AJAX calls.
            fresh = await page.evaluate(REDISCOVER_FULL_JS, headers)
            if not fresh or fresh['maxPage'] <= 1:
                non_empty = [h for h in headers if h]
                print(f"  Section {non_empty[:3]}: single page or no pagination (skipping)")
                continue

            func_name = fresh['funcName']
            view_id = fresh['viewId']
            page_size = fresh['pageSize']
            current_max = fresh['maxPage']

            print(f"  Section: {current_max} pages, pageSize={page_size}, headers: {[h for h in headers if h]}")

            # Ensure on page 1
            try:
                await page.evaluate(
                    f"window['{func_name}']('{view_id}', 0, {page_size}, '', '')"
                )
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(2000)
            except Exception:
                pass

            # Collect page 1
            result = await page.evaluate(FIND_TABLE_ROWS_JS, headers)
            all_rows = result['rows'] if result else []
            print(f"    Page 1: {len(all_rows)} rows")
            if not all_rows:
                continue

            # Navigate pages 2..N
            for pg in range(2, current_max + 1):
                old_fp = await page.evaluate(FINGERPRINT_JS, headers)

                # Re-discover function name (changes after each AJAX)
                cur = await page.evaluate(REDISCOVER_FUNC_JS, headers)
                cf = cur['funcName'] if cur else func_name
                cv = cur['viewId'] if cur else view_id

                offset = (pg - 1) * page_size
                try:
                    await page.evaluate(
                        f"window['{cf}']('{cv}', {offset}, {page_size}, '', '')"
                    )
                except Exception as e:
                    print(f"    Error invoking _pageLoad for page {pg}: {e}")
                    break

                await page.wait_for_load_state("networkidle")
                # Wait for AJAX response to fully update the DOM
                await page.wait_for_timeout(2000)
                try:
                    await page.wait_for_function(WAIT_CHANGE_JS, [headers, old_fp], timeout=8000)
                except Exception:
                    # Extra wait if fingerprint hasn't changed yet
                    await page.wait_for_timeout(3000)

                result = await page.evaluate(FIND_TABLE_ROWS_JS, headers)
                new_rows = result['rows'] if result else []
                added = sum(1 for r in new_rows if r not in all_rows)
                for r in new_rows:
                    if r not in all_rows:
                        all_rows.append(r)
                print(f"    Page {pg}: {len(new_rows)} rows ({added} new, total: {len(all_rows)})")

                # Sliding-window: check if more pages appeared
                um = await page.evaluate("""(eh) => {
                    for (const t of document.querySelectorAll('table')) {
                        let ths = Array.from(t.querySelectorAll('thead th'));
                        if (ths.length === 0) ths = Array.from(t.querySelectorAll('tr:first-child th'));
                        if (eh.every(h => ths.map(th=>th.textContent.trim()).some(th=>th===h))) {
                            let mx=0;
                            t.querySelectorAll('a[onclick*="_pageLoad"]').forEach(a => {
                                const n=parseInt(a.textContent.trim());
                                if(!isNaN(n)&&n>mx) mx=n;
                            });
                            return mx;
                        }
                    }
                    return 0;
                }""", headers)
                if um > current_max:
                    current_max = um

            collected_data.append({'headers': headers, 'rows': all_rows})
            print(f"    Collected {len(all_rows)} total rows for rebuild")

        # ============================================================
        # PHASE 2: REBUILD all tables at once (no more AJAX after this)
        # ============================================================
        if not collected_data:
            return

        # Re-expand all accordion sections (AJAX may have collapsed some)
        await page.evaluate("""() => {
            document.querySelectorAll('.collapse, .panel-collapse, [class*="collapse"]').forEach(el => {
                el.style.display = 'block';
                el.style.height = 'auto';
                el.style.overflow = 'visible';
                el.classList.add('in', 'show');
            });
        }""")
        await page.wait_for_timeout(500)

        print("Rebuilding all tables with collected data...")
        for data in collected_data:
            headers = data['headers']
            rows_html = "".join(data['rows'])
            rebuilt = await page.evaluate("""([expectedHeaders, rowsHtml]) => {
                const tables = document.querySelectorAll('table');
                for (const table of tables) {
                    let ths = Array.from(table.querySelectorAll('thead th'));
                    if (ths.length === 0) ths = Array.from(table.querySelectorAll('tr:first-child th'));
                    const hdr = ths.map(th => th.textContent.trim());
                    if (expectedHeaders.length > 0 &&
                        expectedHeaders.every(h => hdr.some(th => th === h))) {
                        let tbody = table.querySelector('tbody');
                        if (!tbody) { tbody = document.createElement('tbody'); table.appendChild(tbody); }
                        tbody.innerHTML = rowsHtml;
                        return true;
                    }
                }
                return false;
            }""", [headers, rows_html])
            non_empty = [h for h in headers if h]
            print(f"  Rebuilt {non_empty[:3]}... with {len(data['rows'])} rows: {'OK' if rebuilt else 'FAILED'}")

        # Hide any remaining pagination/results elements
        await page.evaluate("""() => {
            document.querySelectorAll('nav').forEach(n => {
                if (n.closest('table')) n.closest('tr').style.display = 'none';
            });
            document.querySelectorAll('.pager, .pagination').forEach(p => p.style.display = 'none');
        }""")

    except Exception as e:
        print(f"Warning: Error expanding pagination: {e}")
        import traceback
        traceback.print_exc()

async def main():
    if len(sys.argv) < 2:
        print("Usage: python riverside_docket_scraper.py <case_number> [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    # Default to headless unless --headful is specified
    is_headless = "--headful" not in sys.argv
    
    login_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=user/login"
    search_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=node/379"

    if not os.environ.get("GEMINI_API_KEY"):
        print("CRITICAL ERROR: GEMINI_API_KEY environment variable is not set.")
        print("The script cannot solve the CAPTCHA without it.")
        print("Please set the environment variable and try again.")
        # Pause to let the user see the message if running in a separate console window
        if not is_headless:
            time.sleep(10) 
        sys.exit(1)

    async with async_playwright() as p:
        # headless=False is required for manual login captcha solving
        print(f"Launching browser for Riverside Superior Court (Headless: {is_headless})...")
        browser = await p.chromium.launch(
            headless=is_headless,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-setuid-sandbox"]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            # Phase 1: Authentication
            print("Navigating to Riverside Login...")
            await page.goto(login_url)
            
            max_login_attempts = 5
            logged_in = False
            
            for attempt in range(1, max_login_attempts + 1):
                print(f"Login attempt {attempt} of {max_login_attempts}...")
                
                # Fill credentials (only if field is empty/visible)
                if await page.locator("#edit-name").is_visible():
                    await page.fill("#edit-name", "Serpiklaw@gmail.com")
                    await page.fill("#edit-pass", "Pikserv321!321!")
                
                print("--- AUTOMATED: Solving CAPTCHA with Gemini ---")
                
                # Solve captcha and click login (solve_image_captcha clicks #edit-submit)
                await solve_image_captcha(page)
                
                # Check for success: username field disappears OR specific success URL/element
                try:
                    # Wait for either the login field to disappear (success) or an error message (failure)
                    # We use a short timeout for each attempt's verification
                    await page.wait_for_selector("#edit-name", state="hidden", timeout=15000)
                    print("Login successful.")
                    logged_in = True
                    break
                except:
                    print(f"Attempt {attempt} failed or timed out. Checking for error messages...")
                    # Check if we are still on the login page. If so, loop continues.
                    # The page often reloads with a new captcha automatically.
                    await page.wait_for_load_state("networkidle")

                    # Check for specific error messages on the page
                    error_selectors = [
                        ".messages--error",
                        ".error-message",
                        ".alert-danger",
                        "[role='alert']",
                        ".form-item--error-message"
                    ]
                    for selector in error_selectors:
                        error_el = page.locator(selector).first
                        if await error_el.count() > 0:
                            error_text = await error_el.inner_text()
                            print(f"  Page error: {error_text.strip()[:200]}")

                            # Check for account lockout - no point retrying
                            if "temporarily blocked" in error_text.lower() or "too many" in error_text.lower():
                                print("ACCOUNT LOCKED: Too many failed attempts. Please wait before retrying.")
                                print("The account lockout typically lasts 15-30 minutes.")
                                sys.exit(1)
            
            if not logged_in:
                print("CRITICAL ERROR: Failed to log in after maximum attempts.")
                sys.exit(1)

            print(f"Current URL: {page.url}")
            print("Proceeding to search...")

            # Phase 2: Case Search
            await page.goto(search_url)
            await page.wait_for_load_state("networkidle")
            
            print(f"Searching for Case: {case_number}")
            
            # Robustly find and fill the case number field using XPath relative to label
            # Structure: label("Case Number") -> sibling div -> input
            try:
                # specific XPath based on analyzed HTML
                await page.fill("//label[contains(., 'Case Number')]/following::input[1]", case_number)
            except Exception as e:
                print(f"Primary XPath failed: {e}. Trying broad text match...")
                # Fallback: Find any input near text "Case Number"
                await page.locator("text=Case Number").locator("xpath=../..").locator("input[type=text]").first.fill(case_number)
            
            # Resolve the math question automatically
            await solve_math_captcha(page)
            
            # Click the search button with increased timeout
            await page.click("#edit-submit", timeout=60000)
            
            # Phase 3: Results and PDF Generation
            print("Waiting for results...")
            case_link = f"a:has-text('{case_number}')"
            try:
                await page.wait_for_selector(case_link, timeout=60000)
            except:
                print(f"Case {case_number} not found in search results.")
                sys.exit(1)

            await page.click(case_link)
            
            # Wait for content rendering
            await page.wait_for_load_state("networkidle")

            # Expand all collapsible accordion sections before looking for pagination.
            # The Riverside court page has collapsed sections (COMPLAINTS/PETITIONS,
            # HEARINGS, COLLECTION HISTORY, DOCUMENTS, CASE LEDGER) whose content
            # is lazy-loaded when expanded. We must click each section header to load
            # its content (including pagination) into the DOM.
            print("Expanding all collapsible sections...")
            section_names = ["COMPLAINTS/PETITIONS", "HEARINGS", "COLLECTION HISTORY", "DOCUMENTS", "CASE LEDGER"]
            for section_name in section_names:
                try:
                    # Find the clickable section header by its text
                    header = page.locator(f"text='{section_name}'").first
                    if await header.count() > 0:
                        await header.click(timeout=5000)
                        print(f"  Expanded: {section_name}")
                        # Wait for AJAX to load section content
                        await page.wait_for_load_state("networkidle")
                        await page.wait_for_timeout(1500)
                    else:
                        print(f"  Not found: {section_name}")
                except Exception as e:
                    print(f"  Could not expand {section_name}: {e}")

            # Give all sections time to fully render
            await page.wait_for_timeout(2000)

            # Expand all paginated sections (Hearings, Documents, etc.)
            print("Expanding all paginated sections...")
            await expand_all_pagination(page)

            # Force screen media so print stylesheets don't collapse sections
            await page.emulate_media(media="screen")

            # Apply CSS fixes: force all accordion/collapse sections visible for PDF
            await page.evaluate("""() => {
                // Force all collapse panels to be visible
                document.querySelectorAll('.collapse, .panel-collapse, .accordion-collapse, [class*="collapse"]').forEach(el => {
                    el.style.display = 'block';
                    el.style.height = 'auto';
                    el.style.overflow = 'visible';
                    el.style.visibility = 'visible';
                    el.style.opacity = '1';
                    el.classList.add('in', 'show');  // Bootstrap 3 'in', Bootstrap 4/5 'show'
                });

                // Force all containers to show content
                const mainContainer = document.querySelector('.main-container');
                if (mainContainer) {
                    mainContainer.style.overflow = 'visible';
                    mainContainer.style.height = 'auto';
                }
                document.body.style.overflow = 'visible';
                document.body.style.height = 'auto';

                // Override any print-specific hiding rules
                const style = document.createElement('style');
                style.textContent = `
                    @media print {
                        .collapse, .panel-collapse, [class*="collapse"] {
                            display: block !important;
                            height: auto !important;
                            overflow: visible !important;
                            visibility: visible !important;
                        }
                        * { overflow: visible !important; }
                    }
                    /* Also force for screen rendering used by PDF */
                    .collapse:not(.navbar-collapse) {
                        display: block !important;
                        height: auto !important;
                    }
                `;
                document.head.appendChild(style);
            }""")

            # Filename format matching docket.py expectation: docket_YYYY.MM.DD.pdf
            output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Generating PDF: {output_filename}...")
            await page.pdf(path=output_filename, format="Letter", print_background=True)
            print(f"Successfully created {output_filename}")

        except Exception as e:
            print(f"Error in Riverside scraper: {e}")
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())