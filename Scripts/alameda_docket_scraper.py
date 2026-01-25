import sys
import asyncio
import time
import re
import os
from playwright.async_api import async_playwright

# Configurations
LOGIN_URL = "https://eportal.alameda.courts.ca.gov/?q=user/login&destination=node/387"

async def solve_math_captcha(page):
    """Automatically extracts, solves, and fills the math captcha if present."""
    try:
        captcha_container = page.locator(".form-item-captcha-response")
        if await captcha_container.count() == 0:
            return True # Not present, that's fine

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
        print("Usage: python alameda_docket_scraper.py <case_number> [--headful] [--headless]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv
    if "--headful" in sys.argv:
        is_headless = False

    print(f"Launching browser for Alameda Superior Court (Headless: {is_headless})...")
    
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
            print("Entering credentials...")
            await page.fill("#edit-name", "Serpiklaw@gmail.com")
            await page.fill("#edit-pass", "Botinok5!")
            
            # Solve Captcha (Optional, if it exists)
            await solve_math_captcha(page)
            
            # Click Login
            print("Clicking Log in...")
            await page.click("#edit-submit[value='Log in']")
            await page.wait_for_load_state("networkidle")
            
            # 2. Perform Search
            print(f"Searching for Case: {case_number}")
            # Wait for the case number input box to appear
            await page.wait_for_selector("[id='59334']", timeout=30000)
            await page.fill("[id='59334']", case_number)
            
            print("Clicking Search...")
            await page.click("#edit-submit[value='Search']")
            await page.wait_for_load_state("networkidle")

            # 3. Find Result
            print("Finding case link...")
            # The user mentioned an example link: <a href="?q=node/364/1481116">TORRES vs BOTANIC TONICS, LLC, et al.</a>
            # We'll try to find a link that contains "vs" or matches the case number if it's there.
            # Usually, the first link in the results table is the case.
            
            # Try to find link by case number or just the first link in the results area
            # A common pattern for these portals is a table with results.
            # We can try to find a link that contains "vs" as it's a common case name pattern.
            case_link = page.locator("a[href*='q=node/364/']").first
            if await case_link.count() == 0:
                # Fallback: look for any link in the content area that might be the case
                case_link = page.locator(".view-content a").first

            if await case_link.count() == 0:
                print(f"Case {case_number} link not found in results.")
                # Print page text to help debug if it fails
                # content = await page.content()
                # print(content)
                sys.exit(1)
            
            case_name = await case_link.inner_text()
            print(f"Opening case: {case_name}")
            await case_link.click()
            await page.wait_for_load_state("networkidle")
            
            # Wait a bit for tabs or content to load
            await page.wait_for_timeout(3000)
            
            # 4. Print to PDF
            output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Generating PDF: {output_filename}...")
            
            # Set to print background and use Letter format
            await page.pdf(path=output_filename, format="Letter", print_background=True)
            
            print(f"Successfully created {output_filename}")

        except Exception as e:
            print(f"Error in Alameda scraper: {e}")
            import traceback
            traceback.print_exc()
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
