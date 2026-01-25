import sys
import asyncio
import time
import re
import os
import datetime
from playwright.async_api import async_playwright
from pypdf import PdfWriter

# Helper function to solve math captcha
async def solve_math_captcha(page):
    """Automatically extracts, solves, and fills the math captcha."""
    print("Attempting to solve math captcha...")
    try:
        # The math problem is usually text inside the parent div or near the input
        # Structure often: <div class="form-item ..."> <label>Math question ...</label> ... <input ...> ... </div>
        # or simply text in the form-item wrapper.
        
        # Locate the container for the captcha response
        captcha_container = page.locator(".form-item-captcha-response")
        
        # If the container isn't found easily, try looking for the label specifically
        if await captcha_container.count() == 0:
            # Fallback: look for the label for the captcha input
            label = page.locator("label[for='edit-captcha-response']")
            if await label.count() > 0:
                container_text = await label.inner_text()
            else:
                # Fallback: get the text of the entire form or a surrounding div
                container_text = await page.locator("form").inner_text()
        else:
            container_text = await captcha_container.inner_text()
            # Sometimes the question is in the parent of the container if the container is just the input wrapper
            # But usually it's close.
            
        print(f"Debug: Captcha context text: '{container_text}'")
        
        # Look for pattern like "1 + 9 =" or "1 + 9" or "Math question: 1 + 9 ="
        # Regex to find two numbers separated by a plus sign
        match = re.search(r'(\d+)\s*\+\s*(\d+)', container_text)
        
        if match:
            num1 = int(match.group(1))
            num2 = int(match.group(2))
            result = num1 + num2
            print(f"Solved math captcha: {num1} + {num2} = {result}")
            
            await page.fill("#edit-captcha-response", str(result))
            return True
        else:
            print("Warning: Could not find math pattern (X + Y) in captcha text.")
            
    except Exception as e:
        print(f"Warning: Could not solve math captcha automatically: {e}")
    return False

async def main():
    if len(sys.argv) < 2:
        print("Usage: python kern_docket_scraper.py <case_number> [--headless]")
        sys.exit(1)

    case_number_input = sys.argv[1]
    is_headless = "--headless" in sys.argv
    
    # Process case number: Remove dashes as requested
    case_number_clean = case_number_input.replace("-", "").strip()
    print(f"Processing Case: {case_number_input} -> Search Term: {case_number_clean}")

    login_url = "https://portal.kern.courts.ca.gov/?q=Login"
    
    async with async_playwright() as p:
        print(f"Launching browser for Kern County (Headless: {is_headless})...")
        browser = await p.chromium.launch(
            headless=is_headless,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-setuid-sandbox"]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            # --- Phase 1: Login ---
            print("Navigating to Kern Login...")
            await page.goto(login_url)
            
            # Fill Credentials
            await page.fill("#edit-name", "aserpik@bordinsemmer.com")
            await page.fill("#edit-pass", "Pikserv123!")
            
            # Solve Captcha
            captcha_solved = await solve_math_captcha(page)
            if not captcha_solved:
                print("Warning: Captcha might not have been solved. Attempting login anyway...")
            
            # Click Login
            print("Clicking Login...")
            await page.click("#edit-submit")
            await page.wait_for_load_state("networkidle")
            
            # Check if login succeeded (Check for "Log out" link or similar, or absence of login form)
            if await page.locator("#edit-name").is_visible():
                print("Login failed (Login form still visible).")
                sys.exit(1)
            print("Login successful.")

            # --- Phase 2: Search ---
            print("Navigating to Case Search...")
            # Click "Case Search" button/link
            await page.click("a[href='/?q=node/393']") # As per user instruction
            await page.wait_for_load_state("networkidle")
            
            print(f"Searching for case number: {case_number_clean}")
            # Input Case Number
            # User provided name="data(168295)". 
            # Note: Parentheses in attribute values usually don't need escaping in CSS selectors if quoted, 
            # but name='data(168295)' might be tricky. Let's try explicit attribute selector.
            await page.fill("input[name='data(168295)']", case_number_clean)
            
            # Click Search
            print("Clicking Search...")
            await page.click("#edit-submit") # Search button has same ID 'edit-submit'
            
            print("Waiting for search results...")
            await page.wait_for_load_state("networkidle")
            
            # Find and Click Case Number
            # User provided example: <a href="?q=node/394/5388770">BCV22101251</a>
            # We look for a link containing the case number text
            case_link_selector = f"a:text-is('{case_number_clean}')"
            
            try:
                await page.wait_for_selector(case_link_selector, timeout=10000)
                print(f"Found case link for {case_number_clean}. Clicking...")
                await page.click(case_link_selector)
            except:
                print(f"Case {case_number_clean} not found in search results.")
                # fallback: try partial match
                try:
                    print("Trying partial match...")
                    await page.click(f"a:has-text('{case_number_clean}')")
                except:
                     print("Failed to find case link.")
                     sys.exit(1)

            await page.wait_for_load_state("networkidle")
            
            # --- Phase 3: Print Pages ---
            # Tabs: Filings, Parties, Documents, Events
            # Selectors: #FV-Filings-Portal, #FV-Parties-Portal, #FV-Documents-Portal, #FV-Events-Portal
            
            tabs = [
                ("Filings", "#FV-Filings-Portal"),
                ("Parties", "#FV-Parties-Portal"),
                ("Documents", "#FV-Documents-Portal"),
                ("Events", "#FV-Events-Portal")
            ]
            
            pdf_files = []
            
            for tab_name, selector in tabs:
                print(f"Processing Tab: {tab_name}...")
                try:
                    # Click the tab
                    await page.click(selector)
                    await page.wait_for_load_state("networkidle")
                    # Give it a moment to render dynamic content
                    await asyncio.sleep(2) 
                    
                    # Print to PDF
                    pdf_filename = f"temp_{tab_name.lower()}.pdf"
                    await page.pdf(path=pdf_filename, format="Letter", print_background=True)
                    pdf_files.append(pdf_filename)
                    print(f"Saved {pdf_filename}")
                    
                except Exception as e:
                    print(f"Error processing tab {tab_name}: {e}")
            
            # --- Phase 4: Merge PDFs ---
            if not pdf_files:
                print("No PDFs were generated.")
                sys.exit(1)
                
            print("Merging PDFs...")
            final_filename = f"docket_{case_number_input}_{datetime.datetime.now().strftime('%Y.%m.%d')}.pdf"
            
            merger = PdfWriter()
            for pdf_file in pdf_files:
                merger.append(pdf_file)
            
            merger.write(final_filename)
            merger.close()
            print(f"Successfully created combined docket: {final_filename}")
            
            # Cleanup temp files
            for pdf_file in pdf_files:
                try:
                    os.remove(pdf_file)
                except:
                    pass

        except Exception as e:
            print(f"Error in Kern scraper: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
