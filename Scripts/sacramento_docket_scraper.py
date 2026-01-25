import sys
import asyncio
import time
import re
import os
import shutil
from playwright.async_api import async_playwright
from pypdf import PdfWriter

# Configurations
LOGIN_URL = "https://prod-portal-sacramento-ca.journaltech.com/public-portal/?q=Login"

async def solve_math_captcha(page):
    """Automatically extracts, solves, and fills the math captcha."""
    try:
        # Structure often used in these portals:
        # <div class="form-item form-type-textfield form-item-captcha-response">
        #   <label>Math question like "1 + 1 =" <span class="form-required" ...>...</span></label>
        #   <input ...>
        # </div>
        # We need the text from the parent or label.
        
        captcha_container = page.locator(".form-item-captcha-response")
        if await captcha_container.count() == 0:
            print("Warning: Captcha container not found.")
            return False

        container_text = await captcha_container.inner_text()
        print(f"Debug: Captcha container text: '{container_text}'")
        
        # Look for pattern like "1 + 9 ="
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
        print("Usage: python sacramento_docket_scraper.py <case_number> [--headful] [--headless]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv
    if "--headful" in sys.argv:
        is_headless = False

    print(f"Launching browser for Sacramento Superior Court (Headless: {is_headless})...")
    
    async with async_playwright() as p:
        browser = await p.chromium.launch(
            headless=is_headless,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-setuid-sandbox"]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            # 1. Login
            print("Navigating to Login...")
            await page.goto(LOGIN_URL)
            await page.wait_for_load_state("networkidle")

            # Fill Credentials
            await page.fill("#edit-name", "serpiklaw@gmail.com")
            await page.fill("#edit-pass", "pIKSERV123!")
            
            # Solve Captcha
            print("Solving CAPTCHA...")
            if not await solve_math_captcha(page):
                print("Failed to solve captcha automatically. Please check logic.")
                # If strictly headless, we might fail here. If headful, maybe user intervenes?
                # For now, we assume success or fail.
            
            # Click Login
            await page.click("#edit-submit")
            await page.wait_for_load_state("networkidle")
            
            # Check if login succeeded (Searches button exists)
            if await page.locator("a[href='?q=node/408']").count() == 0:
                print("Login failed or 'Searches' button not found.")
                # It might be that captcha failed.
                sys.exit(1)
            print("Login successful.")

            # 2. Navigate to Search
            print("Clicking 'Searches'...")
            await page.click("a[href='?q=node/408']")
            await page.wait_for_load_state("networkidle")

            print("Clicking 'Case Number Search'...")
            await page.click("a[href='?q=node/414']")
            await page.wait_for_load_state("networkidle")

            # 3. Perform Search
            print(f"Searching for Case: {case_number}")
            await page.fill("[id='59334']", case_number) # Case Number Box
            
            # Select "Unlimited Civil"
            # Value "CU" based on user prompt
            await page.select_option("[id='59321']", value="CU")
            
            print("Clicking Search...")
            await page.click("#edit-submit[value='Search']")
            await page.wait_for_load_state("networkidle")

            # 4. Find Result
            print("Finding case link...")
            # User said: <a href="?q=node/397/2402803">25CV023819</a>
            # match by exact text of case number
            case_link = page.locator(f"a:text-is('{case_number}')")
            if await case_link.count() == 0:
                print(f"Case {case_number} not found in results.")
                sys.exit(1)
            
            await case_link.click()
            await page.wait_for_load_state("networkidle")
            print("Opened case details.")

            # 5. Print Pages
            pdfs_to_merge = []

            # A. Hearings
            print("Processing 'Hearings'...")
            hearings_tab = page.locator("a.ec-tab-link:has-text('Hearings')")
            if await hearings_tab.is_visible():
                await hearings_tab.click()
                # Wait for content to load? Usually these are AJAX tabs.
                # Inspecting similar portals, usually there's a delay.
                await page.wait_for_timeout(2000) 
                
                pdf_path = "temp_hearings.pdf"
                await page.pdf(path=pdf_path, format="Letter", print_background=True)
                pdfs_to_merge.append(pdf_path)
                print(f"Saved {pdf_path}")
            else:
                print("'Hearings' tab not found.")

            # B. Case Participants
            print("Processing 'Case Participants'...")
            participants_tab = page.locator("a.ec-tab-link:has-text('Case Participants')")
            if await participants_tab.is_visible():
                await participants_tab.click()
                await page.wait_for_timeout(2000)
                
                pdf_path = "temp_participants.pdf"
                await page.pdf(path=pdf_path, format="Letter", print_background=True)
                pdfs_to_merge.append(pdf_path)
                print(f"Saved {pdf_path}")
            else:
                print("'Case Participants' tab not found.")

            # C. Register of Actions
            print("Processing 'Register of Actions'...")
            roa_tab = page.locator("a.ec-tab-link:has-text('Register of Actions')")
            if await roa_tab.is_visible():
                await roa_tab.click()
                await page.wait_for_timeout(2000)
                
                pdf_path = "temp_roa.pdf"
                await page.pdf(path=pdf_path, format="Letter", print_background=True)
                pdfs_to_merge.append(pdf_path)
                print(f"Saved {pdf_path}")
            else:
                print("'Register of Actions' tab not found.")

            # 6. Combine PDFs
            if not pdfs_to_merge:
                print("No PDFs were generated.")
                sys.exit(1)

            final_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Merging into {final_filename}...")
            
            merger = PdfWriter()
            for pdf in pdfs_to_merge:
                merger.append(pdf)
            
            merger.write(final_filename)
            merger.close()
            
            print(f"Successfully created {final_filename}")

            # Cleanup
            for pdf in pdfs_to_merge:
                if os.path.exists(pdf):
                    os.remove(pdf)

        except Exception as e:
            print(f"Error in Sacramento scraper: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
