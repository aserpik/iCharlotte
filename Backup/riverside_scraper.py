import sys
import asyncio
import time
import re
from playwright.async_api import async_playwright

async def solve_math_captcha(page):
    """Automatically extracts, solves, and fills the math captcha."""
    try:
        # Standard Drupal math captcha label selector
        label_text = await page.inner_text("label[for='edit-captcha-response']")
        # Use regex to find all numbers in the text (e.g., "5 + 3 =")
        numbers = re.findall(r'\d+', label_text)
        if len(numbers) >= 2:
            result = int(numbers[0]) + int(numbers[1])
            await page.fill("#edit-captcha-response", str(result))
            print(f"Solved math captcha: {numbers[0]} + {numbers[1]} = {result}")
            return True
    except Exception as e:
        print(f"Warning: Could not solve math captcha automatically: {e}")
    return False

async def main():
    if len(sys.argv) < 2:
        print("Usage: python riverside_scraper.py <case_number>")
        sys.exit(1)

    case_number = sys.argv[1]
    login_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=user/login"
    search_url = "https://epublic-access.riverside.courts.ca.gov/public-portal/?q=node/379"

    async with async_playwright() as p:
        # headless=False is required for manual login captcha solving
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()
        page = await context.new_page()

        # Phase 1: Authentication
        print("Navigating to Riverside Login...")
        await page.goto(login_url)
        await page.fill("#edit-name", "Serpiklaw@gmail.com")
        await page.fill("#edit-pass", "Pikserv123!123!123!")
        
        print("--- ACTION REQUIRED: Solve the image CAPTCHA and click Log In ---")
        # Wait for user to login; the "Log out" text appears once successful
        await page.wait_for_selector("text=Log out", timeout=120000)
        print("Login detected. Proceeding to search...")

        # Phase 2: Case Search
        await page.goto(search_url)
        print(f"Searching for Case: {case_number}")
        await page.fill("input[name='case_number']", case_number)
        
        # Resolve the math question automatically
        await solve_math_captcha(page)
        
        # Click the search button (ID confirmed as edit-submit)
        await page.click("#edit-submit")
        
        # Phase 3: Results and PDF Generation
        print("Waiting for results...")
        case_link = f"a:has-text('{case_number}')"
        await page.wait_for_selector(case_link)
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

        output_filename = f"docket_{time.strftime('%Y.%m.%d')}.pdf"
        print(f"Generating PDF: {output_filename}...")
        await page.pdf(path=output_filename, format="Letter", print_background=True)
        
        await browser.close()
        print("Riverside download complete.")

if __name__ == "__main__":
    asyncio.run(main())