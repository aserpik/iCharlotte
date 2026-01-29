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
    Expands all paginated tables by collecting all data from each page.
    Modifies the DOM to show all rows in each table before PDF generation.
    """
    try:
        # Find all tables that have associated pagination
        # We'll process each table section separately
        tables = page.locator("table")
        table_count = await tables.count()

        if table_count == 0:
            print("No tables found on page.")
            return

        print(f"Found {table_count} table(s). Checking for pagination...")

        # Process tables in reverse order to avoid index shifting issues
        for table_idx in range(table_count - 1, -1, -1):
            table = tables.nth(table_idx)

            # Find the section/container that holds this table
            # Look for a parent with a pager nearby
            section = await table.evaluate("""el => {
                // Walk up to find a container that has both table and pager
                let parent = el.parentElement;
                for (let i = 0; i < 10 && parent; i++) {
                    if (parent.querySelector('.pager') || parent.querySelector('.pagination')) {
                        return parent.className || parent.id || 'found';
                    }
                    parent = parent.parentElement;
                }
                return null;
            }""")

            if not section:
                continue

            print(f"Processing table {table_idx + 1} with pagination...")

            # Collect all rows from all pages for this table
            all_rows_html = []
            header_html = None
            page_num = 1
            max_pages = 50  # Safety limit

            while page_num <= max_pages:
                # Re-fetch table reference as DOM may have changed
                current_table = page.locator("table").nth(table_idx)

                if await current_table.count() == 0:
                    break

                # Get header on first page
                if header_html is None:
                    thead = current_table.locator("thead")
                    if await thead.count() > 0:
                        header_html = await thead.evaluate("el => el.outerHTML")

                # Get current table body rows
                tbody = current_table.locator("tbody")
                if await tbody.count() > 0:
                    rows = tbody.locator("tr")
                    row_count = await rows.count()

                    for r in range(row_count):
                        row = rows.nth(r)
                        row_html = await row.evaluate("el => el.outerHTML")
                        # Avoid duplicate rows
                        if row_html not in all_rows_html:
                            all_rows_html.append(row_html)

                    print(f"  Page {page_num}: collected {row_count} rows (total: {len(all_rows_html)})")

                # Find the pager associated with this table section
                # Look for next page link
                pager = page.locator(".pager").first
                next_link = pager.locator("a[title='Go to next page'], .pager-next a").first

                if await next_link.count() == 0:
                    print(f"  No more pages (stopped at page {page_num})")
                    break

                # Check if next link is actually clickable
                is_visible = await next_link.is_visible()
                if not is_visible:
                    print(f"  Next link not visible (stopped at page {page_num})")
                    break

                # Click next and wait for content to load
                await next_link.click()
                await page.wait_for_load_state("networkidle")
                await page.wait_for_timeout(800)  # Wait for table to update

                page_num += 1

            # If we collected rows from multiple pages, rebuild the table
            if page_num > 1 and len(all_rows_html) > 0:
                print(f"  Rebuilding table with all {len(all_rows_html)} rows...")

                # Join all rows into single HTML string
                all_rows_joined = "".join(all_rows_html)
                total_rows = len(all_rows_html)

                # Inject all collected rows into the table using evaluate with arguments
                # This safely passes the HTML without string escaping issues
                await page.evaluate("""([tableIdx, rowsHtml, totalRows]) => {
                    const table = document.querySelectorAll('table')[tableIdx];
                    if (!table) return;

                    let tbody = table.querySelector('tbody');
                    if (!tbody) {
                        tbody = document.createElement('tbody');
                        table.appendChild(tbody);
                    }

                    // Replace tbody content with all collected rows
                    tbody.innerHTML = rowsHtml;

                    // Hide pagination since we now show all data
                    const pagers = document.querySelectorAll('.pager, .pagination');
                    pagers.forEach(p => p.style.display = 'none');

                    // Update any "Results X-Y of Z" text
                    const resultTexts = document.querySelectorAll('[class*="result"], [class*="count"]');
                    resultTexts.forEach(el => {
                        if (el.textContent.includes('Results') || el.textContent.includes('of')) {
                            el.textContent = 'Showing all ' + totalRows + ' results';
                        }
                    });
                }""", [table_idx, all_rows_joined, total_rows])

                print(f"  Table rebuilt successfully with {len(all_rows_html)} rows")

        # Final cleanup: hide all pagination controls
        await page.evaluate("""() => {
            document.querySelectorAll('.pager, .pagination, .pager-nav').forEach(p => {
                p.style.display = 'none';
            });
        }""")

    except Exception as e:
        print(f"Warning: Error expanding pagination: {e}")
        import traceback
        traceback.print_exc()
        # Continue anyway - we'll get at least the first page

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
                    await page.fill("#edit-pass", "Pikserv123!123!123!")
                
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

            # Expand all paginated sections (Hearings, Documents, etc.)
            print("Expanding all paginated sections...")
            await expand_all_pagination(page)

            # Apply CSS fixes for clean PDF formatting
            await page.evaluate("""() => {
                const mainContainer = document.querySelector('.main-container');
                if (mainContainer) {
                    mainContainer.style.overflow = 'visible';
                    mainContainer.style.height = 'auto';
                }
                document.body.style.overflow = 'visible';
                document.body.style.height = 'auto';
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