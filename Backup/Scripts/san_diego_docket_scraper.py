import sys
import asyncio
import time
import os
from playwright.async_api import async_playwright

async def main():
    if len(sys.argv) < 2:
        print("Usage: python san_diego_docket_scraper.py <case_number> [--headless]")
        sys.exit(1)

    case_number = sys.argv[1]
    # Default to headless unless --headful is specified (or keep it always headless if that's the goal)
    is_headless = "--headful" not in sys.argv
    
    start_url = "https://odyroa.sdcourt.ca.gov/Cases"

    async with async_playwright() as p:
        print(f"Launching browser for San Diego Superior Court (Headless: {is_headless})...")
        browser = await p.chromium.launch(headless=is_headless)
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        )
        page = await context.new_page()

        try:
            # Step 1: Go to the URL
            print(f"Navigating to {start_url}...")
            await page.goto(start_url)
            
            # Step 2: Click "I have read..." button
            # Selector based on the text provided by user
            print("Accepting terms...")
            # Using partial text match or exact class/text combo
            await page.click("text=I have read, understood, and agree to abide by the above terms.")
            
            # Wait for the search page to load
            await page.wait_for_selector("#CaseNumber", state="visible")

            # Step 3: Enter case number
            print(f"Entering case number: {case_number}")
            await page.fill("#CaseNumber", case_number)
            
            # Step 4: Click Search
            print("Clicking Search...")
            await page.click("input[value='Search']")
            
            # Step 5: Click Printable ROA
            # This might be on the resulting page. We need to wait for it.
            print("Waiting for results and 'Printable ROA' link...")
            try:
                # Wait specifically for the Printable ROA link to appear
                # The user provided example HTML: <a class="btn btn-primary" href="/cases/print/...">Printable ROA</a>
                await page.wait_for_selector("text=Printable ROA", timeout=60000)
                await page.click("text=Printable ROA")
            except Exception as e:
                print(f"Error: 'Printable ROA' button not found. Case might not exist or search failed. Details: {e}")
                # Take a screenshot for debugging if headless
                if is_headless:
                    await page.screenshot(path="san_diego_error.png")
                sys.exit(1)

            # Step 6: Print to PDF
            # The Printable ROA button likely navigates to a new page or opens a print view.
            # We should wait for navigation or the new content to load.
            await page.wait_for_load_state("networkidle")
            
            # Filename format matching docket.py expectation: docket_YYYY.MM.DD.pdf
            output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Generating PDF: {output_filename}...")
            
            # Ensure background graphics are printed (often needed for proper formatting)
            await page.pdf(path=output_filename, format="Letter", print_background=True)
            print(f"Successfully created {output_filename}")

        except Exception as e:
            print(f"Error in San Diego scraper: {e}")
            sys.exit(1)
        finally:
            # Clean up debug screenshot if it exists
            if os.path.exists("san_diego_error.png"):
                try:
                    os.remove("san_diego_error.png")
                except:
                    pass
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
