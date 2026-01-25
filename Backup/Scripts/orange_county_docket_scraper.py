import sys
import asyncio
import time
import os
import re
import json
import base64
import google.generativeai as genai
from playwright.async_api import async_playwright

async def solve_recaptcha(page):
    """
    Solves Google reCAPTCHA v2 using Gemini (Audio First, fallback to Vision).
    Mimics 'Buster' extension logic by attempting audio challenge.
    """
    debug_log = open("orange_debug.log", "a")
    def log_debug(msg):
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        debug_log.write(f"[{timestamp}] {msg}\n")
        debug_log.flush()
        print(msg)

    try:
        log_debug("Starting reCAPTCHA solver (Audio First)...")
        
        # Wait for any iframe to load
        try:
            await page.wait_for_selector("iframe", timeout=10000)
        except:
            log_debug("Timeout waiting for iframes.")

        # 1. Find the reCAPTCHA frame
        recaptcha_frame = None
        try:
            frame_element = await page.wait_for_selector("iframe[src*='recaptcha/api2/anchor']", timeout=20000)
            recaptcha_frame = await frame_element.content_frame()
        except Exception as e:
            log_debug(f"Error waiting for anchor frame selector: {e}")
        
        if not recaptcha_frame:
            log_debug("Frame not found via selector. Listing all frames...")
            for f in page.frames:
                if "google.com/recaptcha" in f.url and "anchor" in f.url:
                    recaptcha_frame = f
                    break

        if not recaptcha_frame:
            log_debug("Could not find reCAPTCHA anchor frame.")
            return False

        log_debug(f"Found anchor frame: {recaptcha_frame.url}")
        
        # Ensure the checkbox is visible before clicking
        try:
            await recaptcha_frame.wait_for_selector("#recaptcha-anchor", timeout=10000)
            log_debug("Clicking checkbox...")
            await recaptcha_frame.click("#recaptcha-anchor")
        except Exception as e:
            log_debug(f"Error clicking anchor: {e}")
            return False

        # Check if solved immediately
        await asyncio.sleep(3)
        try:
            is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
            if is_checked == "true":
                log_debug("reCAPTCHA solved immediately (no challenge).")
                return True
        except:
            pass

        log_debug("Waiting for challenge frame...")
        challenge_frame = None
        for _ in range(15):
            all_frames = page.frames
            for f in all_frames:
                if "google.com/recaptcha" in f.url and "bframe" in f.url:
                    try:
                        frame_el = await f.frame_element()
                        if frame_el and await frame_el.is_visible():
                            challenge_frame = f
                            break
                    except:
                        continue
            if challenge_frame:
                break
            await asyncio.sleep(1)

        if not challenge_frame:
            log_debug("Could not resolve visible challenge frame object.")
            return False

        # --- AUDIO SOLVER STRATEGY ---
        log_debug("Attempting Audio Solver (Buster Logic)...")
        
        api_key = os.environ.get("GEMINI_API_KEY")
        if not api_key:
            log_debug("GEMINI_API_KEY not set.")
            return False
        genai.configure(api_key=api_key)

        async def get_gemini_audio_response(audio_path):
            models = ["gemini-1.5-flash", "gemini-3-flash-preview"] # 1.5 Flash is reliable for audio
            for model_name in models:
                try:
                    log_debug(f"Processing audio with {model_name}...")
                    model = genai.GenerativeModel(model_name)
                    sample_file = genai.upload_file(path=audio_path, display_name=f"Audio_Captcha_{time.time()}")
                    
                    # Wait for processing state if necessary (usually fast for small audio)
                    while sample_file.state.name == "PROCESSING":
                        time.sleep(1)
                        sample_file = genai.get_file(sample_file.name)

                    prompt = "Listen to this audio numbers and return ONLY the numbers you hear as a sequence of digits. No other text."
                    response = model.generate_content([sample_file, prompt])
                    text = response.text.strip()
                    # Extract digits only
                    digits = "".join(filter(str.isdigit, text))
                    return digits
                except Exception as e:
                    log_debug(f"Error with {model_name}: {e}")
                    continue
            return None

        # Click Audio Button
        try:
            audio_button = challenge_frame.locator("#recaptcha-audio-button")
            if await audio_button.is_visible():
                log_debug("Clicking Audio Button...")
                await audio_button.click()
                await asyncio.sleep(2)
                
                # Check for "Try again later" block
                content = await challenge_frame.content()
                if "Try again later" in content or "dostad" in content: # 'dostad' is often part of the block msg class
                    log_debug("Audio challenge blocked ('Try again later'). Falling back to Vision...")
                else:
                    # Look for audio source
                    log_debug("Looking for audio source...")
                    audio_src = None
                    try:
                        # Try download link first
                        dl_link = challenge_frame.locator(".rc-audiochallenge-tdownload-link")
                        if await dl_link.count() > 0:
                            audio_src = await dl_link.get_attribute("href")
                        
                        if not audio_src:
                            # Try audio tag
                            audio_tag = challenge_frame.locator("#audio-source")
                            if await audio_tag.count() > 0:
                                audio_src = await audio_tag.get_attribute("src")
                    except:
                        pass
                    
                    if audio_src:
                        log_debug(f"Found audio source: {audio_src[:30]}...")
                        # Download audio using fetch in browser context to keep cookies
                        audio_data_b64 = await challenge_frame.evaluate(f"""
                            async () => {{
                                const response = await fetch('{audio_src}');
                                const buffer = await response.arrayBuffer();
                                return btoa(String.fromCharCode(...new Uint8Array(buffer)));
                            }}
                        """)
                        
                        audio_file_path = f"captcha_audio_{time.time()}.mp3"
                        with open(audio_file_path, "wb") as f:
                            f.write(base64.b64decode(audio_data_b64))
                        
                        log_debug("Audio downloaded. Solving...")
                        digits = await get_gemini_audio_response(audio_file_path)
                        
                        # Cleanup audio
                        try:
                            os.remove(audio_file_path)
                        except:
                            pass

                        if digits:
                            log_debug(f"Solved Digits: {digits}")
                            await challenge_frame.fill("#audio-response", digits)
                            await challenge_frame.click("#recaptcha-verify-button")
                            await asyncio.sleep(2)
                            
                            try:
                                is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
                                if is_checked == "true":
                                    log_debug("Audio reCAPTCHA solved!")
                                    return True
                            except:
                                pass
                            
                            # Check for error
                            if await challenge_frame.locator(".rc-audiochallenge-error-message").is_visible():
                                log_debug("Audio verification failed. Retrying or Fallback...")
                        else:
                            log_debug("Could not transcribe audio.")
                    else:
                        log_debug("No audio source found (or UI changed).")

        except Exception as e:
            log_debug(f"Audio solver exception: {e}")
        
        # --- FALLBACK TO VISION SOLVER (Original Logic) ---
        log_debug("Falling back to Vision Solver...")
        
        # Ensure we are on Image challenge (click Image button if we are on Audio)
        try:
            image_button = challenge_frame.locator("#recaptcha-image-button")
            if await image_button.is_visible():
                 log_debug("Switching back to Image challenge...")
                 await image_button.click()
                 await asyncio.sleep(1)
        except:
            pass

        async def get_gemini_vision_response(image_path, prompt):
            models = ["gemini-3-flash-preview", "gemini-3-pro-preview", "gemini-1.5-flash"]
            for model_name in models:
                try:
                    log_debug(f"Attempting generation with {model_name}...")
                    model = genai.GenerativeModel(model_name)
                    sample_file = genai.upload_file(path=image_path, display_name=f"reCAPTCHA_{time.time()}")
                    response = model.generate_content([sample_file, prompt])
                    return response.text.strip()
                except Exception as e:
                    log_debug(f"Error with {model_name}: {e}")
                    if "429" in str(e) or "quota" in str(e).lower():
                        continue
                    continue
            return None

        # Dynamic solving loop (Vision)
        max_rounds = 20
        for round_idx in range(max_rounds):
            log_debug(f"--- Vision Round {round_idx + 1} ---")
            
            try:
                is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
                if is_checked == "true":
                    log_debug("reCAPTCHA solved!")
                    return True
            except:
                pass
            
            # 1. Get Instruction
            instruction_text = ""
            try:
                if await challenge_frame.locator(".rc-imageselect-desc-wrapper").count() > 0:
                    instruction_text = await challenge_frame.inner_text(".rc-imageselect-desc-wrapper")
                elif await challenge_frame.locator(".rc-imageselect-desc").count() > 0:
                    instruction_text = await challenge_frame.inner_text(".rc-imageselect-desc")
                elif await challenge_frame.locator("strong").count() > 0:
                    instruction_text = await challenge_frame.inner_text("strong")
                else:
                    instruction_text = await challenge_frame.inner_text("body")
                instruction_text = instruction_text.replace("\n", " ").strip()
                log_debug(f"Instruction: {instruction_text}")
            except Exception as e:
                log_debug(f"Error reading instruction: {e}")

            # 2. Screenshot Target
            target_selector = "#rc-imageselect-target"
            if await challenge_frame.locator(target_selector).count() == 0:
                log_debug("Image target not found. Checking if solved...")
                try:
                    is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
                    if is_checked == "true":
                        return True
                except:
                    pass
                log_debug("Target gone. Retrying loop...")
                continue

            image_path = f"recaptcha_round_{round_idx}.png"
            try:
                await challenge_frame.locator(target_selector).screenshot(path=image_path, timeout=10000)
            except Exception as e:
                log_debug(f"Screenshot failed: {e}")
                await page.screenshot(path=image_path)

            # 3. Check grid type
            is_4x4 = await challenge_frame.locator("table.rc-imageselect-table-44").count() > 0
            grid_size = 16 if is_4x4 else 9
            
            # 4. Ask Gemini
            prompt = f"""
            This is a reCAPTCHA challenge.
            Instruction: "{instruction_text}"
            Grid size: {grid_size} squares (3x3 or 4x4).
            
            Task: Identify the squares that match the instruction.
            - If the instruction says "Click verify once there are none left", select ALL matching instances.
            - If it's a standard static challenge, select the matching squares.
            - If NO squares match the instruction, return an empty list [].
            
            Return ONLY a JSON list of numbers (1-{grid_size}).
            Top-left is 1, reading left-to-right, row-by-row.
            """
            
            try:
                text = await get_gemini_vision_response(image_path, prompt)
                if not text:
                    log_debug("All Gemini models failed to provide a response.")
                    indices = []
                else:
                    log_debug(f"Gemini response: {text}")
                    match = re.search(r'\[.*?\]', text, re.DOTALL)
                    indices = []
                    if match:
                        indices = json.loads(match.group(0))
            except Exception as e:
                log_debug(f"Error parsing Gemini response: {e}")
                indices = []

            log_debug(f"Indices to click: {indices}")

            # 5. Click or Verify
            if not indices:
                log_debug("No matches found. Clicking Verify/Skip...")
                verify_btn = challenge_frame.locator("#recaptcha-verify-button")
                try:
                    await verify_btn.click(force=True, timeout=5000)
                except:
                    pass
                await asyncio.sleep(2)
            else:
                log_debug(f"Clicking {len(indices)} tiles...")
                tiles = challenge_frame.locator(".rc-imageselect-tile")
                for idx in indices:
                    try:
                        tile = tiles.nth(idx - 1)
                        await tile.click(force=True, timeout=3000)
                    except:
                        pass
                    if "none left" in instruction_text.lower():
                        await asyncio.sleep(1.0) 
                    else:
                        await asyncio.sleep(0.1)
                
                if "none left" not in instruction_text.lower():
                    log_debug("Static challenge likely. Clicking Verify...")
                    verify_btn = challenge_frame.locator("#recaptcha-verify-button")
                    try:
                        await verify_btn.click(force=True, timeout=5000)
                    except:
                        pass
                    await asyncio.sleep(2)
                else:
                    await asyncio.sleep(1.5)

    except Exception as e:
        log_debug(f"Error in reCAPTCHA solver: {e}")
    finally:
        # Cleanup temporary PNG files
        log_debug("Cleaning up temporary reCAPTCHA files...")
        for file in os.listdir("."):
            if (file.startswith("recaptcha_round_") and file.endswith(".png")) or \
               file in ["debug_no_captcha_frame.png", "debug_no_challenge_frame.png"] or \
               (file.startswith("debug_gone_") and file.endswith(".png")):
                try:
                    os.remove(file)
                except:
                    pass
    
    return False

async def main():
    if len(sys.argv) < 3:
        print("Usage: python orange_county_docket_scraper.py <case_number> <filing_year> [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    filing_year = sys.argv[2]
    # Default to headful unless --headless is specified
    is_headless = "--headless" in sys.argv
    
    url = "https://civilwebshopping.occourts.org/Login.do"

    async with async_playwright() as p:
        print(f"Launching browser for Orange County (Headless: {is_headless})...")
        browser = await p.chromium.launch(
            headless=is_headless,
            args=["--disable-blink-features=AutomationControlled", "--no-sandbox", "--disable-setuid-sandbox"]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            viewport={'width': 1920, 'height': 1080},
            is_mobile=False,
            has_touch=False,
            permissions=['geolocation'],
            device_scale_factor=1
        )
        # Stealth: Undefine navigator.webdriver
        await context.add_init_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")
        
        page = await context.new_page()

        try:
            print(f"Navigating to {url}...")
            await page.goto(url)

            # 1. Accept Terms
            print("Accepting terms...")
            await page.click("input[value='Accept Terms']")
            await page.wait_for_selector("#caseNumber", timeout=10000)

            # 2. Enter Data
            print(f"Entering Case Number: {case_number}")
            await page.fill("#caseNumber", case_number)
            
            print(f"Entering Year: {filing_year}")
            await page.fill("#caseYear", filing_year)
            await page.evaluate("document.getElementById('caseYear').blur()")

            # 3. Solve reCAPTCHA
            print("Solving reCAPTCHA...")
            if await solve_recaptcha(page):
                print("reCAPTCHA solved successfully.")
            else:
                print("Failed to solve reCAPTCHA. Aborting.")
                # We exit here to avoid clicking 'Search' with an open captcha popup
                sys.exit(1)
            
            # 4. Search
            print("Clicking Search...")
            # Using force=True to bypass potential overlays if captcha is closed but still animating
            await page.click("#action", force=True)
            
            # 5. Print Case
            print("Waiting for 'Print Case' button...")
            try:
                # Wait for search results or print button
                await page.wait_for_selector("#printCase", timeout=30000)
                
                # Handle print action
                # We assume clicking it opens a new page or navigates
                async with context.expect_page() as new_page_info:
                    await page.click("#printCase")
                    # If it doesn't open a new page, this will timeout, so we wrap in try/except or logic?
                    # But printCase usually pops up.
                
                print_page = await new_page_info.value
                await print_page.wait_for_load_state("networkidle")
                
                output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
                print(f"Generating PDF: {output_filename}...")
                await print_page.pdf(path=output_filename, format="Letter", print_background=True)
                print("Success.")

            except Exception as e:
                # Fallback: if no new page, maybe we are on the page?
                print(f"Note: Popup not detected, checking current page... ({e})")
                if await page.locator("#printCase").count() > 0:
                     # Just print current page if we are still here? 
                     # But usually printCase changes the view. 
                     # If the user says "click on the print case button... then print the resulting page",
                     # we must assume the page changes or opens.
                     pass
                sys.exit(1)

        except Exception as e:
            print(f"Error: {e}")
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
