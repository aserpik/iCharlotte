import sys
import asyncio
import time
import os
from playwright.async_api import async_playwright

async def main():
    if len(sys.argv) < 2:
        print("Usage: python santa_clara_docket_scraper.py <case_number> [--headless] [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    # Default to headless for PDF support, allow override
    is_headless = "--headful" not in sys.argv
    
    login_url = "https://portal.scscourt.org/login"
    EMAIL = "serpiklaw@gmail.com"
    PASSWORD = "Pikserv123!"

    async with async_playwright() as p:
        print(f"Launching browser for Santa Clara Superior Court (Headless: {is_headless})...")
        browser = await p.chromium.launch(headless=is_headless)
        context = await browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            # Step 1: Login
            print(f"Navigating to {login_url}...")
            await page.goto(login_url, wait_until="networkidle")
            
            print("Inputting credentials...")
            # Using selectors from user's provided elements
            await page.wait_for_selector("#email", timeout=10000)
            await page.fill("#email", EMAIL)
            await page.fill("#password", PASSWORD)
            
            print("Clicking Log In...")
            # Button element: <button class="btn btn-primary btn-custom w-md waves-effect waves-light" type="submit">
            await page.click('button[type="submit"]')
            
            # Wait for navigation after login
            print("Waiting for dashboard to load...")
            await page.wait_for_load_state("networkidle")
            
            # Step 2: Search Case
            # Input element: <input id="search-input" ... placeholder="CASE NUMBER / PERSON NAME / BUSINESS NAME" ...>
            print(f"Searching for case number: {case_number}...")
            search_input_selector = "#search-input"
            await page.wait_for_selector(search_input_selector, timeout=30000)
            await page.click(search_input_selector)
            await page.fill(search_input_selector, case_number)
            await page.keyboard.press("Enter")
            
            # Step 3: Wait for results and click printable button
            print("Waiting for case details page...")
            # The page change after Enter might take a moment
            await page.wait_for_load_state("networkidle")
            
            # Click the PRINTABLE tab first to make the button visible
            print("Clicking PRINTABLE tab...")
            printable_tab_selector = 'a[data-target="#full"]'
            try:
                await page.wait_for_selector(printable_tab_selector, timeout=30000)
                await page.click(printable_tab_selector)
                
                # Wait for the data tables to actually populate. 
                # These portals use DataTables with AJAX, so we check for the 'processing' indicator or presence of rows.
                print("Waiting for data tables to populate...")
                # Wait for the PARTIES table to have at least one row that isn't the "Loading" or "No data" placeholder if possible,
                # or just wait for the network to settle and the DOM to be stable.
                await page.wait_for_selector("#printableView table tbody tr", timeout=30000)
                # Small additional delay for Angular's digest cycle and rendering
                await page.wait_for_timeout(3000)
            except Exception as e:
                print(f"Warning: Content might not be fully loaded: {e}")

            # Printable button element provided by user:
            # <button type="button" class="btn btn-primary pull-right" onclick="window.print();">
            printable_button_selector = 'button.btn-primary.pull-right:has-text("Print")'
            
            try:
                await page.wait_for_selector(printable_button_selector, state="visible", timeout=30000)
            except:
                print("Case details page did not load as expected or 'Print' button not found.")
                raise Exception("Print button not found.")

            # Step 4: Handle the print dialogue (Save to PDF)
            print("Generating PDF (Auto-saving)...")
            date_str = time.strftime("%Y.%m.%d")
            output_filename = f"docket_{case_number}_{date_str}.pdf"
            
            # To capture the printable version, we emulate print media
            await page.emulate_media(media="print")
            
            # Programmatic PDF generation is the automated equivalent of "Clicking Save" in the print dialog
            if is_headless:
                await page.pdf(path=output_filename, format="Letter", print_background=True)
                print(f"Successfully saved docket to {output_filename}")
            else:
                # In headful mode, page.pdf() is not supported.
                # We can try to use screenshot as a fallback or inform the user.
                print("Note: Automated PDF saving via 'page.pdf' requires headless mode.")
                print("Taking a full-page screenshot as a fallback...")
                output_filename = f"docket_{case_number}_{date_str}.png"
                await page.screenshot(path=output_filename, full_page=True)
                print(f"Saved screenshot to {output_filename}")

        except Exception as e:
            print(f"An error occurred: {e}")
            await page.screenshot(path=f"error_santa_clara_{case_number}.png")
            # For debugging, also save the HTML
            with open(f"error_santa_clara_{case_number}.html", "w", encoding="utf-8") as f:
                f.write(await page.content())
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())