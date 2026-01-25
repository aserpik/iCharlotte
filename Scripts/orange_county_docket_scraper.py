import sys
import asyncio
import time
import os
import re
import json
import random
from google import genai
from google.genai import types
from playwright.async_api import async_playwright

async def solve_recaptcha(page):
    """
    Solves Google reCAPTCHA v2 using Gemini Vision (image grid challenges only).
    """
    debug_log = open("orange_debug.log", "a")
    def log_debug(msg):
        timestamp = time.strftime('%Y-%m-%d %H:%M:%S')
        debug_log.write(f"[{timestamp}] {msg}\n")
        debug_log.flush()
        print(msg)

    try:
        log_debug("Starting reCAPTCHA solver (Vision Only - Gemini)...")

        # Initialize Gemini
        gemini_key = os.environ.get("GEMINI_API_KEY")
        if not gemini_key:
            log_debug("GEMINI_API_KEY not set.")
            return False
        client = genai.Client(api_key=gemini_key)

        # Wait for grecaptcha API to be ready
        log_debug("Waiting for grecaptcha API...")
        try:
            await page.wait_for_function("typeof grecaptcha !== 'undefined'", timeout=10000)
            log_debug("grecaptcha API detected")
        except:
            log_debug("grecaptcha API not found after 10s")

        # Check if reCAPTCHA div exists and try to trigger it
        try:
            recaptcha_div = await page.query_selector(".g-recaptcha")
            if recaptcha_div:
                log_debug("Found g-recaptcha div, scrolling into view...")
                await recaptcha_div.scroll_into_view_if_needed()
                await asyncio.sleep(1)

                # Check current state
                has_iframe = await page.evaluate("document.querySelector('.g-recaptcha iframe') !== null")
                log_debug(f"reCAPTCHA already has iframe: {has_iframe}")

                if not has_iframe:
                    # Debug: Check what's in the div and grecaptcha object
                    log_debug("Debugging reCAPTCHA setup...")
                    try:
                        debug_info = await page.evaluate("""
                            () => {
                                const div = document.querySelector('.g-recaptcha');
                                const info = {
                                    divExists: !!div,
                                    divHTML: div ? div.outerHTML.substring(0, 500) : null,
                                    divAttributes: div ? Array.from(div.attributes).map(a => a.name + '=' + a.value) : [],
                                    grecaptchaExists: typeof grecaptcha !== 'undefined',
                                    grecaptchaMethods: typeof grecaptcha !== 'undefined' ? Object.keys(grecaptcha) : [],
                                    grecaptchaEnterprise: typeof grecaptcha !== 'undefined' && typeof grecaptcha.enterprise !== 'undefined',
                                };
                                return info;
                            }
                        """)
                        log_debug(f"Debug info: {debug_info}")
                    except Exception as e:
                        log_debug(f"Debug error: {e}")

                    # Try different render approaches
                    log_debug("Attempting to trigger reCAPTCHA render...")
                    try:
                        result = await page.evaluate("""
                            () => {
                                const div = document.querySelector('.g-recaptcha');
                                if (!div) return 'no div';

                                // Get sitekey from various sources
                                let sitekey = div.getAttribute('data-sitekey');
                                if (!sitekey) {
                                    // Try to find it in scripts
                                    const scripts = document.querySelectorAll('script');
                                    for (const s of scripts) {
                                        const match = s.textContent.match(/sitekey['":\s]+['"]([^'"]+)['"]/i);
                                        if (match) {
                                            sitekey = match[1];
                                            break;
                                        }
                                    }
                                }
                                if (!sitekey) {
                                    // Check for data-sitekey in any element
                                    const el = document.querySelector('[data-sitekey]');
                                    if (el) sitekey = el.getAttribute('data-sitekey');
                                }

                                if (!sitekey) return 'no sitekey found anywhere';

                                if (typeof grecaptcha === 'undefined') return 'no grecaptcha';

                                // Try enterprise API first
                                if (grecaptcha.enterprise && grecaptcha.enterprise.render) {
                                    try {
                                        grecaptcha.enterprise.render(div, {sitekey: sitekey});
                                        return 'rendered via enterprise';
                                    } catch (e) {
                                        return 'enterprise render error: ' + e.message;
                                    }
                                }

                                // Try standard API
                                if (grecaptcha.render) {
                                    try {
                                        grecaptcha.render(div, {sitekey: sitekey});
                                        return 'rendered';
                                    } catch (e) {
                                        return 'render error: ' + e.message;
                                    }
                                }

                                return 'no render function available, methods: ' + Object.keys(grecaptcha).join(',');
                            }
                        """)
                        log_debug(f"Render result: {result}")
                    except Exception as e:
                        log_debug(f"grecaptcha.render attempt: {e}")

                await asyncio.sleep(3)
        except Exception as e:
            log_debug(f"Error checking recaptcha div: {e}")

        # Wait for any iframe to load
        try:
            await page.wait_for_selector("iframe", timeout=15000)
        except:
            log_debug("Timeout waiting for iframes.")

        # 1. Find the reCAPTCHA anchor frame
        recaptcha_frame = None
        try:
            frame_element = await page.wait_for_selector("iframe[src*='recaptcha/api2/anchor']", timeout=15000)
            recaptcha_frame = await frame_element.content_frame()
        except Exception as e:
            log_debug(f"Error waiting for anchor frame selector: {e}")

        if not recaptcha_frame:
            log_debug("Frame not found via selector. Listing all frames...")
            for f in page.frames:
                log_debug(f"  Frame: {f.url[:100] if f.url else '(no url)'}")
                if "google.com/recaptcha" in f.url and "anchor" in f.url:
                    recaptcha_frame = f
                    break

        if not recaptcha_frame:
            log_debug("Could not find reCAPTCHA anchor frame.")
            # Save debug screenshot
            await page.screenshot(path="debug_no_recaptcha_frame.png")
            log_debug("Debug screenshot saved to debug_no_recaptcha_frame.png")
            # Also check page content for clues
            content = await page.content()
            has_recaptcha_div = "g-recaptcha" in content
            has_recaptcha_script = "recaptcha" in content.lower()
            log_debug(f"Page has g-recaptcha div: {has_recaptcha_div}")
            log_debug(f"Page mentions recaptcha: {has_recaptcha_script}")
            return False

        log_debug(f"Found anchor frame: {recaptcha_frame.url}")

        # Click the checkbox with human-like delay
        try:
            await recaptcha_frame.wait_for_selector("#recaptcha-anchor", timeout=10000)
            await asyncio.sleep(0.5 + (time.time() % 1))  # Random delay
            log_debug("Clicking checkbox...")
            await recaptcha_frame.click("#recaptcha-anchor")
        except Exception as e:
            log_debug(f"Error clicking anchor: {e}")
            return False

        # Check if solved immediately (no challenge)
        await asyncio.sleep(3)
        try:
            is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
            if is_checked == "true":
                log_debug("reCAPTCHA solved immediately (no challenge).")
                return True
        except:
            pass

        # Wait for challenge frame
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

        # Check for "Try again later" block immediately
        content = await challenge_frame.content()
        if "Try again later" in content:
            log_debug("BLOCKED: 'Try again later' detected. Google has flagged this session.")
            return False

        # Helper function to upload image to Gemini with proper MIME type
        async def get_gemini_vision_response(image_path, prompt):
            models = ["gemini-2.0-flash", "gemini-1.5-flash"]
            for model_name in models:
                try:
                    log_debug(f"Attempting vision with {model_name}...")

                    # Read file and upload with explicit mime_type
                    with open(image_path, "rb") as f:
                        file_data = f.read()

                    # Use inline data with explicit mime type instead of file upload
                    image_part = types.Part.from_bytes(
                        data=file_data,
                        mime_type="image/png"
                    )

                    response = client.models.generate_content(
                        model=model_name,
                        contents=[image_part, prompt]
                    )
                    return response.text.strip()
                except Exception as e:
                    log_debug(f"Error with {model_name}: {e}")
                    continue
            return None

        # --- VISION SOLVER ---
        log_debug("Starting Vision Solver...")

        max_rounds = 15
        for round_idx in range(max_rounds):
            log_debug(f"--- Vision Round {round_idx + 1} ---")

            # Check if already solved
            try:
                is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
                if is_checked == "true":
                    log_debug("reCAPTCHA SOLVED!")
                    return True
            except:
                pass

            # Check for block message
            content = await challenge_frame.content()
            if "Try again later" in content:
                log_debug("BLOCKED: 'Try again later' detected.")
                return False

            # Get the instruction text
            instruction_text = ""
            try:
                # Try multiple selectors for instruction
                for selector in [".rc-imageselect-desc-wrapper", ".rc-imageselect-desc", ".rc-imageselect-instructions"]:
                    if await challenge_frame.locator(selector).count() > 0:
                        instruction_text = await challenge_frame.inner_text(selector)
                        break

                if not instruction_text:
                    # Fallback to strong tag
                    if await challenge_frame.locator("strong").count() > 0:
                        instruction_text = await challenge_frame.inner_text("strong")

                instruction_text = instruction_text.replace("\n", " ").strip()
                log_debug(f"Instruction: {instruction_text}")
            except Exception as e:
                log_debug(f"Error reading instruction: {e}")
                continue

            # Check if image grid is present
            target_selector = "#rc-imageselect-target"
            if await challenge_frame.locator(target_selector).count() == 0:
                log_debug("No image grid found, waiting...")
                await asyncio.sleep(1)
                continue

            # Take screenshot of the image grid
            image_path = f"recaptcha_round_{round_idx}.png"
            try:
                await challenge_frame.locator(target_selector).screenshot(path=image_path, timeout=10000)
            except Exception as e:
                log_debug(f"Screenshot failed: {e}")
                await asyncio.sleep(1)
                continue

            # Determine grid size (3x3 = 9 tiles, 4x4 = 16 tiles)
            is_4x4 = await challenge_frame.locator("table.rc-imageselect-table-44").count() > 0
            is_dynamic = "none left" in instruction_text.lower() or "keep clicking" in instruction_text.lower()
            grid_size = 16 if is_4x4 else 9
            grid_desc = "4x4" if is_4x4 else "3x3"

            log_debug(f"Grid: {grid_desc} ({grid_size} tiles), Dynamic: {is_dynamic}")

            # Build a detailed prompt for Gemini
            prompt = f"""You are solving a reCAPTCHA image challenge.

INSTRUCTION FROM CAPTCHA: "{instruction_text}"

The image shows a {grid_desc} grid of tiles numbered as follows:
{"1  2  3  4" if is_4x4 else "1  2  3"}
{"5  6  7  8" if is_4x4 else "4  5  6"}
{"9  10 11 12" if is_4x4 else "7  8  9"}
{"13 14 15 16" if is_4x4 else ""}

Analyze each tile carefully. Which tiles match the instruction?

IMPORTANT RULES:
- Return ONLY a JSON array of tile numbers, e.g. [1, 3, 7]
- If NO tiles match, return an empty array: []
- Numbers must be between 1 and {grid_size}
- Be precise - only select tiles that CLEARLY match the target object
- Look for the MAIN subject described (e.g., "traffic lights" means the actual light fixture, not just poles)

Your response (JSON array only):"""

            # Get Gemini's response
            text = await get_gemini_vision_response(image_path, prompt)
            log_debug(f"Gemini response: {text}")

            # Parse the response
            indices = []
            try:
                if text:
                    # Find JSON array in response
                    match = re.search(r'\[[\d,\s]*\]', text)
                    if match:
                        indices = json.loads(match.group(0))
                        # Filter to valid range
                        indices = [i for i in indices if isinstance(i, int) and 1 <= i <= grid_size]
            except Exception as e:
                log_debug(f"Parse error: {e}")
                indices = []

            log_debug(f"Tiles to click: {indices}")

            # Click the tiles
            if indices:
                tiles = challenge_frame.locator(".rc-imageselect-tile")
                for idx in indices:
                    try:
                        await asyncio.sleep(0.2 + (time.time() % 0.3))  # Human-like delay
                        await tiles.nth(idx - 1).click(force=True, timeout=3000)
                        log_debug(f"Clicked tile {idx}")
                    except Exception as e:
                        log_debug(f"Failed to click tile {idx}: {e}")

                    # For dynamic challenges, wait for tile to reload
                    if is_dynamic:
                        await asyncio.sleep(1.5)

            # Wait a moment then click verify (unless it's a dynamic "click until none left" challenge)
            await asyncio.sleep(0.5)

            # For dynamic challenges, check if new tiles appeared
            if is_dynamic and indices:
                # Wait for possible tile refresh
                await asyncio.sleep(2)
                # Continue to next round to re-analyze
                continue

            # Click verify button
            try:
                verify_btn = challenge_frame.locator("#recaptcha-verify-button")
                if await verify_btn.is_visible():
                    log_debug("Clicking Verify...")
                    await verify_btn.click(force=True)
                    await asyncio.sleep(2)
            except Exception as e:
                log_debug(f"Error clicking verify: {e}")

            # Check if we got an error message (wrong selection)
            try:
                error_msg = challenge_frame.locator(".rc-imageselect-error-select-more, .rc-imageselect-error-dynamic-more")
                if await error_msg.is_visible():
                    log_debug("Error: Need to select more tiles")
                    continue
            except:
                pass

        # Final check
        try:
            is_checked = await recaptcha_frame.evaluate("document.querySelector('#recaptcha-anchor').getAttribute('aria-checked')")
            if is_checked == "true":
                log_debug("reCAPTCHA SOLVED!")
                return True
        except:
            pass

        log_debug("Failed to solve reCAPTCHA after maximum rounds.")
        return False

    except Exception as e:
        log_debug(f"Error in reCAPTCHA solver: {e}")
        import traceback
        log_debug(traceback.format_exc())
        return False
    finally:
        debug_log.close()
        # Cleanup temp files
        for file in os.listdir("."):
            if file.startswith("recaptcha_round_"):
                try:
                    os.remove(file)
                except:
                    pass

async def main():
    if len(sys.argv) < 3:
        print("Usage: python orange_county_docket_scraper.py <case_number> <filing_year> [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    filing_year = sys.argv[2]
    # Default to headless unless --headful is specified
    is_headless = "--headful" not in sys.argv

    url = "https://civilwebshopping.occourts.org/Login.do"

    async with async_playwright() as p:
        print(f"Launching browser for Orange County (Headless: True - forced)...")

        browser = await p.chromium.launch(
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--disable-dev-shm-usage",
                "--no-sandbox",
                "--disable-setuid-sandbox",
                "--disable-accelerated-2d-canvas",
                "--disable-gpu",
            ]
        )

        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
            viewport={'width': 1920, 'height': 1080},
            is_mobile=False,
            has_touch=False,
            locale='en-US',
            timezone_id='America/Los_Angeles',
            device_scale_factor=1,
            java_script_enabled=True,
            bypass_csp=False,
            extra_http_headers={
                'Accept-Language': 'en-US,en;q=0.9',
                'Accept-Encoding': 'gzip, deflate, br',
                'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8',
                'sec-ch-ua': '"Chromium";v="122", "Not(A:Brand";v="24", "Google Chrome";v="122"',
                'sec-ch-ua-mobile': '?0',
                'sec-ch-ua-platform': '"Windows"',
                'Sec-Fetch-Dest': 'document',
                'Sec-Fetch-Mode': 'navigate',
                'Sec-Fetch-Site': 'none',
                'Sec-Fetch-User': '?1',
                'Upgrade-Insecure-Requests': '1',
            }
        )

        # Comprehensive stealth scripts to avoid bot detection
        await context.add_init_script("""
            // Remove webdriver property
            Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
            delete navigator.__proto__.webdriver;

            // Mock plugins (important for headless detection)
            Object.defineProperty(navigator, 'plugins', {
                get: () => {
                    const plugins = [
                        {
                            name: 'Chrome PDF Plugin',
                            description: 'Portable Document Format',
                            filename: 'internal-pdf-viewer',
                            length: 1,
                            0: {type: 'application/x-google-chrome-pdf', suffixes: 'pdf', description: 'Portable Document Format'}
                        },
                        {
                            name: 'Chrome PDF Viewer',
                            description: '',
                            filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai',
                            length: 1,
                            0: {type: 'application/pdf', suffixes: 'pdf', description: ''}
                        },
                        {
                            name: 'Native Client',
                            description: '',
                            filename: 'internal-nacl-plugin',
                            length: 2,
                            0: {type: 'application/x-nacl', suffixes: '', description: 'Native Client Executable'},
                            1: {type: 'application/x-pnacl', suffixes: '', description: 'Portable Native Client Executable'}
                        }
                    ];
                    plugins.item = (i) => plugins[i];
                    plugins.namedItem = (name) => plugins.find(p => p.name === name);
                    plugins.refresh = () => {};
                    return plugins;
                }
            });

            // Mock mimeTypes
            Object.defineProperty(navigator, 'mimeTypes', {
                get: () => {
                    const mimeTypes = [
                        {type: 'application/pdf', suffixes: 'pdf', description: '', enabledPlugin: navigator.plugins[1]},
                        {type: 'application/x-google-chrome-pdf', suffixes: 'pdf', description: 'Portable Document Format', enabledPlugin: navigator.plugins[0]},
                        {type: 'application/x-nacl', suffixes: '', description: 'Native Client Executable', enabledPlugin: navigator.plugins[2]},
                        {type: 'application/x-pnacl', suffixes: '', description: 'Portable Native Client Executable', enabledPlugin: navigator.plugins[2]}
                    ];
                    mimeTypes.item = (i) => mimeTypes[i];
                    mimeTypes.namedItem = (name) => mimeTypes.find(m => m.type === name);
                    return mimeTypes;
                }
            });

            // Mock languages
            Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
            Object.defineProperty(navigator, 'language', {get: () => 'en-US'});

            // Mock hardware concurrency (number of CPU cores)
            Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});

            // Mock device memory
            Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});

            // Mock max touch points
            Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});

            // Mock connection
            Object.defineProperty(navigator, 'connection', {
                get: () => ({
                    effectiveType: '4g',
                    rtt: 50,
                    downlink: 10,
                    saveData: false
                })
            });

            // Mock permissions
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );

            // Mock chrome runtime
            window.chrome = {
                runtime: {
                    connect: () => {},
                    sendMessage: () => {},
                    onMessage: {addListener: () => {}, removeListener: () => {}},
                    onConnect: {addListener: () => {}, removeListener: () => {}},
                    id: 'mhjfbmdgcfjbbpaeojofohoefgiehjai'
                },
                loadTimes: () => ({
                    commitLoadTime: Date.now() / 1000 - Math.random() * 5,
                    connectionInfo: 'h2',
                    finishDocumentLoadTime: Date.now() / 1000 - Math.random() * 2,
                    finishLoadTime: Date.now() / 1000 - Math.random(),
                    firstPaintAfterLoadTime: 0,
                    firstPaintTime: Date.now() / 1000 - Math.random() * 3,
                    navigationType: 'Other',
                    npnNegotiatedProtocol: 'h2',
                    requestTime: Date.now() / 1000 - Math.random() * 6,
                    startLoadTime: Date.now() / 1000 - Math.random() * 5,
                    wasAlternateProtocolAvailable: false,
                    wasFetchedViaSpdy: true,
                    wasNpnNegotiated: true
                }),
                csi: () => ({
                    onloadT: Date.now(),
                    pageT: Date.now() - Math.random() * 10000,
                    startE: Date.now() - Math.random() * 10000,
                    tran: 15
                }),
                app: {
                    isInstalled: false,
                    InstallState: {DISABLED: 'disabled', INSTALLED: 'installed', NOT_INSTALLED: 'not_installed'},
                    RunningState: {CANNOT_RUN: 'cannot_run', READY_TO_RUN: 'ready_to_run', RUNNING: 'running'}
                }
            };

            // Fix for headless detection via iframe contentWindow
            const originalContentWindow = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'contentWindow');
            Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', {
                get: function() {
                    const win = originalContentWindow.get.call(this);
                    if (win) {
                        try {
                            Object.defineProperty(win.navigator, 'webdriver', {get: () => undefined});
                        } catch (e) {}
                    }
                    return win;
                }
            });

            // Mock WebGL vendor and renderer
            const getParameter = WebGLRenderingContext.prototype.getParameter;
            WebGLRenderingContext.prototype.getParameter = function(parameter) {
                if (parameter === 37445) return 'Google Inc. (NVIDIA)';
                if (parameter === 37446) return 'ANGLE (NVIDIA, NVIDIA GeForce GTX 1080 Direct3D11 vs_5_0 ps_5_0, D3D11)';
                return getParameter.call(this, parameter);
            };

            const getParameter2 = WebGL2RenderingContext.prototype.getParameter;
            WebGL2RenderingContext.prototype.getParameter = function(parameter) {
                if (parameter === 37445) return 'Google Inc. (NVIDIA)';
                if (parameter === 37446) return 'ANGLE (NVIDIA, NVIDIA GeForce GTX 1080 Direct3D11 vs_5_0 ps_5_0, D3D11)';
                return getParameter2.call(this, parameter);
            };

            // Prevent detection via toString
            const nativeToStringFunctionString = Error.toString().replace(/Error/g, 'toString');
            const oldToString = Function.prototype.toString;
            function newToString() {
                if (this === window.navigator.permissions.query) {
                    return 'function query() { [native code] }';
                }
                return oldToString.call(this);
            }
            Function.prototype.toString = newToString;
        """)
        
        page = await context.new_page()

        try:
            print(f"Navigating to {url}...")
            await page.goto(url, wait_until="networkidle")

            # Random delay to appear human
            await asyncio.sleep(1 + random.random() * 2)

            # Move mouse randomly to simulate human behavior
            await page.mouse.move(random.randint(100, 500), random.randint(100, 400))
            await asyncio.sleep(0.3 + random.random() * 0.5)

            # 1. Accept Terms
            print("Accepting terms...")
            await page.click("input[value='Accept Terms']")
            await page.wait_for_selector("#caseNumber", timeout=10000)

            # Another small delay
            await asyncio.sleep(0.5 + random.random())

            # 2. Enter Data
            print(f"Entering Case Number: {case_number}")
            await page.fill("#caseNumber", case_number)

            print(f"Entering Year: {filing_year}")
            await page.fill("#caseYear", filing_year)
            await page.evaluate("document.getElementById('caseYear').blur()")

            # Wait for reCAPTCHA to potentially load
            print("Waiting for page to settle...")
            await asyncio.sleep(3)

            # Take debug screenshot
            await page.screenshot(path="debug_before_recaptcha.png")
            print("Debug screenshot saved to debug_before_recaptcha.png")

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
