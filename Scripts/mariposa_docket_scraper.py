import sys
import asyncio
import time
import os
from playwright.async_api import async_playwright

async def main():
    if len(sys.argv) < 2:
        print("Usage: python mariposa_docket_scraper.py <case_number> [--headless]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv # Default to True unless --headful is specified
    
    url = "https://portal-camariposa.tylertech.cloud/Portal/Home/Dashboard/26"

    async with async_playwright() as p:
        print(f"Launching browser for Mariposa County Superior Court (Headless: {is_headless})...")
        browser = await p.chromium.launch(headless=is_headless)
        context = await browser.new_context(
            accept_downloads=True,
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        )
        page = await context.new_page()

        try:
            print(f"Navigating to {url}...")
            await page.goto(url, wait_until="networkidle", timeout=60000)

            # (2) Select "Case Number" from drop down
            print("Selecting 'Case Number' search type...")
            await page.select_option("select#cboHSSearchBy", value="CaseNumber")

            # (3) Input case_number
            print(f"Inputting case number: {case_number}")
            await page.fill("input#SearchCriteria_SearchValue", case_number)

            # (4) Input Date From
            print("Inputting 'Date From' (01/01/2025)...")
            await page.fill("input#SearchCriteria_DateFrom", "01/01/2025")

            # (5) Input Date To
            print("Inputting 'Date To' (12/01/2026)...")
            await page.fill("input#SearchCriteria_DateTo", "12/01/2026")

            # (6) Click "I'm not a robot"
            print("Attempting to click ReCaptcha...")
            try:
                # ReCaptcha has two main parts: the anchor (checkbox) and the bframe (challenge)
                anchor_iframe_selector = 'iframe[title*="reCAPTCHA"]'
                await page.wait_for_selector(anchor_iframe_selector, timeout=15000)
                
                anchor_frame = None
                for f in page.frames:
                    if "api2/anchor" in f.url:
                        anchor_frame = f
                        break
                
                if anchor_frame:
                    print("Clicking 'I'm not a robot' checkbox...")
                    await anchor_frame.click("#recaptcha-anchor")
                    
                    if not is_headless:
                        print("\n" + "="*60)
                        print("ACTION REQUIRED: PLEASE SOLVE THE RECAPTCHA CHALLENGE NOW.")
                        print("The script will wait up to 5 minutes for you to finish.")
                        print("Once you see the green checkmark, the script will continue.")
                        print("="*60 + "\n")
                        
                        # Reactive loop to check for completion
                        solved = False
                        for _ in range(300): # 300 seconds = 5 minutes
                            try:
                                if await anchor_frame.is_visible(".recaptcha-checkbox-checked"):
                                    print("\nReCaptcha checkmark detected! Proceeding...")
                                    solved = True
                                    break
                            except:
                                pass
                            
                            # Also check if a challenge frame appeared to provide feedback
                            challenge_frame = None
                            for f in page.frames:
                                if "api2/bframe" in f.url:
                                    challenge_frame = f
                                    break
                            
                            if challenge_frame and await challenge_frame.is_visible("#rc-imageselect"):
                                if _ % 10 == 0:
                                    print("Image challenge detected... waiting for you to solve it...")
                            
                            await asyncio.sleep(1)
                        
                        if not solved:
                            print("Timed out waiting for ReCaptcha checkmark. Attempting to proceed anyway...")
                    else:
                        print("WARNING: Running in headless mode. ReCaptcha will likely block this request.")
                        await asyncio.sleep(5)
                else:
                    print("ReCaptcha anchor frame not found. Trying direct click as fallback...")
                    await page.click("#recaptcha-anchor", timeout=5000)
            except Exception as e:
                print(f"ReCaptcha step encountered an issue: {e}. Attempting to proceed to Submit.")

            # (7) Click Submit
            print("Clicking Submit...")
            # Sometimes pressing Enter on the input is more reliable than clicking a blocked button
            await page.focus("input#SearchCriteria_SearchValue")
            await page.keyboard.press("Enter")
            
            # Fallback to click if Enter didn't trigger navigation
            await asyncio.sleep(2)
            if "Dashboard" in page.url:
                 print("Enter key didn't trigger navigation, trying direct button click...")
                 await page.click("input#btnHSSubmit", force=True)

            # (8) Click the first case number link
            print("Searching for case link...")
            case_link_selector = "a.caseLink"
            try:
                # Wait for the case link to be present
                await page.wait_for_selector(case_link_selector, timeout=30000)
                print("Clicking first case link...")
                # We want the first one if multiple appear
                await page.locator(case_link_selector).first.click()
                await page.wait_for_load_state("networkidle")
            except Exception as e:
                print(f"Could not find or click case link: {e}")
                # Take a screenshot to see what's on the page (e.g., "No results found" or ReCaptcha error)
                await page.screenshot(path=f"search_results_error_{case_number}.png")
                # Check if maybe we are already on the page or it failed
                if "CaseDetail" not in page.url:
                    raise Exception("Failed to navigate to Case Detail.")

            # (9) Print the resulting page to PDF
            print("Generating PDF...")
            date_str = time.strftime("%Y.%m.%d")
            output_filename = f"docket_{case_number}_{date_str}.pdf"
            
            await page.pdf(path=output_filename, format="Letter", print_background=True)
            print(f"Successfully created {output_filename}")

        except Exception as e:
            print(f"An error occurred: {e}")
            try:
                if not page.is_closed():
                    await page.screenshot(path=f"error_mariposa_{case_number}.png")
            except:
                pass
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
