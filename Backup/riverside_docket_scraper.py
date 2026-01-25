import sys
import asyncio
import time
import re
import os
import google.generativeai as genai
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
            
        genai.configure(api_key=api_key)
        # Switch to gemini-2.0-flash-exp for potentially better OCR on distorted images
        model = genai.GenerativeModel("gemini-2.0-flash-exp")
        sample_file = genai.upload_file(path=image_path, display_name="Captcha Image")
        
        # Tailored prompt
        if is_image_captcha:
            prompt = """
            Return ONLY the alphanumeric characters shown in this distorted image. 
            Be extremely precise. Look closely at each character:
            - Distinguish carefully between numbers and letters (e.g., '8' vs 'g', '5' vs 'S', '0' vs 'O', '2' vs 'Z').
            - Maintain exact case sensitivity (uppercase vs lowercase).
            Return just the characters, nothing else.
            """
        else:
            prompt = "Solve this math problem (e.g., '1 + 5 ='). Return ONLY the result number."
        
        response = model.generate_content([sample_file, prompt])
        captcha_text = response.text.strip()
        
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

async def main():
    if len(sys.argv) < 2:
        print("Usage: python riverside_docket_scraper.py <case_number>")
        sys.exit(1)

    case_number = sys.argv[1]
    login_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=user/login"
    search_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=node/379"

    if not os.environ.get("GEMINI_API_KEY"):
        print("CRITICAL ERROR: GEMINI_API_KEY environment variable is not set.")
        print("The script cannot solve the CAPTCHA without it.")
        print("Please set the environment variable and try again.")
        # Pause to let the user see the message if running in a separate console window
        time.sleep(10) 
        sys.exit(1)

    async with async_playwright() as p:
        # headless=False is required for manual login captcha solving
        print("Launching browser for Riverside Superior Court...")
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
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
                    print(f"Attempt {attempt} failed or timed out. Checking for CAPTCHA error...")
                    # Check if we are still on the login page. If so, loop continues.
                    # The page often reloads with a new captcha automatically.
                    await page.wait_for_load_state("networkidle")
            
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
            
            # Click the search button
            await page.click("#edit-submit")
            
            # Phase 3: Results and PDF Generation
            print("Waiting for results...")
            case_link = f"a:has-text('{case_number}')"
            try:
                await page.wait_for_selector(case_link, timeout=30000)
            except:
                print(f"Case {case_number} not found in search results.")
                sys.exit(1)

            await page.click(case_link)
            
            # Wait for content rendering
            await page.wait_for_load_state("networkidle")
            
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
            output_filename = f"docket_{time.strftime('%Y.%m.%d')}.pdf"
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