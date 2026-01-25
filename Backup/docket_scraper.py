import sys
import asyncio
import time
from playwright.async_api import async_playwright

async def main():
    if len(sys.argv) < 2:
        print("Usage: python docket_scraper.py <case_number>")
        sys.exit(1)

    case_number = sys.argv[1]
    # Direct link to the application inside the iframe
    url = "https://www.lacourt.ca.gov/casesummary/v2web3/?casetype=civil"

    async with async_playwright() as p:
        print("Launching browser...")
        browser = await p.chromium.launch(headless=True, args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-setuid-sandbox"])
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36"
        )
        page = await context.new_page()
        
        print(f"Navigating to {url}...")
        try:
            await page.goto(url, timeout=60000, wait_until="domcontentloaded")
            print(f"Current URL: {page.url}")
            print(f"Page title: {await page.title()}")
            
            # Debug: save initial content
            with open("initial_page.html", "w", encoding='utf-8') as f:
                f.write(await page.content())

            # Wait for the input field to be ready
            print("Waiting for input field...")
            await page.wait_for_selector("#txtCaseNumber", state="visible", timeout=60000)

            # Fill case number
            print(f"Searching for case: {case_number}")
            await page.fill("#txtCaseNumber", case_number)

            # Click search
            print("Clicking search...")
            await page.click("#submit1")
            
            # Wait for results. 
            # Strategy: Wait for network idle to ensure data fetch is done.
            # Also explicit wait for spinner to vanish if it appears.
            try:
                 await page.wait_for_selector("#waitSpinner", state="visible", timeout=5000)
                 await page.wait_for_selector("#waitSpinner", state="hidden", timeout=60000)
            except Exception as e:
                 # If spinner was too fast or didn't appear, just continue
                 pass
            
            # Additional safety wait for rendering
            await page.wait_for_load_state("networkidle")
            
            # Apply CSS fixes for printing based on user's analysis of the scrollbar location
            print("Applying CSS fixes for PDF printing...")
            await page.evaluate("""
                () => {
                    const contentDiv = document.querySelector('.content');
                    if (contentDiv) {
                        contentDiv.style.height = 'auto';
                        contentDiv.style.overflow = 'visible';
                        // Force width if necessary to avoid horizontal scroll clipping
                        contentDiv.style.width = '100%';
                    }
                    document.body.style.overflow = 'visible';
                    document.body.style.height = 'auto';
                    
                    // Hide non-essential elements for cleaner print if possible (optional)
                    // document.querySelectorAll('.header, .footer, .sidebar').forEach(el => el.style.display = 'none');
                }
            """)

            output_filename = time.strftime("docket_%Y.%m.%d.pdf")
            print(f"DEBUG: Page content before PDF generation:\n{(await page.content())[:1000]}...") # Log partial content
            print(f"Saving PDF to {output_filename}...")
            
            try:
                await page.pdf(path=output_filename, format="Letter")
            except Exception as pdf_error:
                print(f"ERROR: Failed to generate PDF: {pdf_error}")
                raise pdf_error
        except Exception as e:
            print(f"Error occurred: {e}")
            try:
                await page.screenshot(path="error_screenshot.png")
                with open("error_page.html", "w", encoding='utf-8') as f:
                    f.write(await page.content())
                print("Saved error_screenshot.png and error_page.html for debugging.")
            except Exception as inner_e:
                print(f"Could not save debug info: {inner_e}")
            raise e
        
        await browser.close()
        print("Done.")

if __name__ == "__main__":
    asyncio.run(main())
