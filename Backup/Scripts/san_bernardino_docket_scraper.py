import sys
import asyncio
import time
import os
from playwright.async_api import async_playwright

async def main():
    if len(sys.argv) < 2:
        print("Usage: python san_bernardino_docket_scraper.py <case_number> [--headless]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or True # Default to True
    
    login_url = "https://cap.sb-court.org/login"
    
    # Credentials
    EMAIL = "serpiklaw@gmail.com"
    PASSWORD = "Pikserv123!123!"

    async with async_playwright() as p:
        print(f"Launching browser for San Bernardino Superior Court (Headless: {is_headless})...")
        # Headless=True is standard for file downloads/interception
        browser = await p.chromium.launch(headless=is_headless)
        context = await browser.new_context(accept_downloads=True)
        page = await context.new_page()

        try:
            # --- Phase 1: Login ---
            print(f"Navigating to {login_url}...")
            await page.goto(login_url)

            await page.wait_for_selector("input#email", state="visible")
            
            print("Logging in...")
            await page.fill("input#email", EMAIL)
            await page.fill("input#password", PASSWORD)
            await page.press("input#password", "Enter")

            # Wait for dashboard or search input
            try:
                await page.wait_for_selector("input#search-input", state="visible", timeout=30000)
                print("Login successful.")
            except:
                print("Login might have failed or timed out.")
                sys.exit(1)

            # --- Phase 2: Search ---
            print(f"Searching for Case: {case_number}")
            await page.fill("input#search-input", case_number)
            await page.press("input#search-input", "Enter")

            # Wait for the Print button
            print_button_selector = 'button[ng-click="vm.printCaseDetails();"]'
            print("Waiting for Print button...")
            await page.wait_for_selector(print_button_selector, state="visible", timeout=30000)

            # --- Phase 3: Intercept PDF or Extract from Viewer ---
            print("Clicking Print button...")
            
            # Start waiting for a potential download event just in case
            try:
                async with page.expect_download(timeout=5000) as download_info:
                    await page.click(print_button_selector)
                    download = await download_info.value
                    output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
                    await download.save_as(output_filename)
                    print(f"Download detected! Saved to {output_filename}")
                    sys.exit(0)
            except Exception as e:
                # This will catch the TimeoutError if no download starts
                print(f"No direct download detected (Error: {e}). Waiting for PDF Viewer...")

            # Wait for PDF.js viewer to mount
            await page.wait_for_selector("div#viewer.pdfViewer", timeout=30000)
            print("PDF Viewer loaded. Attempting to extract raw PDF data...")

            # Give it a second to initialize the PDF document
            await page.wait_for_timeout(3000)

            # Extract data using PDF.js internal API
            # We access the PDFViewerApplication global object usually present with PDF.js
            pdf_data = await page.evaluate("""async () => {
                try {
                    // Check if PDFViewerApplication is available
                    if (window.PDFViewerApplication && window.PDFViewerApplication.pdfDocument) {
                        const data = await window.PDFViewerApplication.pdfDocument.getData();
                        return Array.from(data);
                    }
                    // Fallback: Check for embedded blob URL if possible (less reliable)
                    return null;
                } catch (e) {
                    return null;
                }
            }""")

            if pdf_data:
                # Use a more unique name to avoid permission issues if a file is locked
                date_str = time.strftime("%Y.%m.%d")
                output_filename = f"docket_{case_number}_{date_str}.pdf"
                
                print(f"Extracted {len(pdf_data)} bytes from PDF Viewer.")
                
                # Convert list back to bytes
                byte_data = bytes(pdf_data)
                try:
                    with open(output_filename, "wb") as f:
                        f.write(byte_data)
                    print(f"Successfully created {output_filename}")
                    
                    # Try to rename it to the expected name, or just let docket.py handle the pattern
                    # docket.py looks for "docket_*.pdf" and picks the newest one.
                except Exception as write_err:
                    print(f"Error writing to {output_filename}: {write_err}")
                    sys.exit(1)
            else:
                print("Failed to extract data via PDFViewerApplication.")
                # Fallback: Try to print the page, but force all pages to render
                print("Attempting fallback print (forcing rendering)...")
                
                # Force render all pages in PDF.js
                await page.evaluate("""() => {
                    if (window.PDFViewerApplication) {
                        window.PDFViewerApplication.pdfViewer.scrollMode = 1; // Vertical scrolling
                        window.PDFViewerApplication.pdfViewer.spreadMode = 0; // No spreads
                    }
                }""")
                
                output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
                await page.pdf(path=output_filename, format="Letter")
                print(f"Saved fallback PDF snapshot to {output_filename}")

                
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())