import sys
import asyncio
import time
import os
import re
import json
import random
import urllib.request
import urllib.parse
from playwright.async_api import async_playwright
from dotenv import load_dotenv
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '.env')
load_dotenv(_env_path, override=True)

async def solve_recaptcha(page):
    """Solve reCAPTCHA v2 using 2Captcha token-based API."""

    api_key = os.environ.get("TWOCAPTCHA_API_KEY")
    if not api_key:
        print("TWOCAPTCHA_API_KEY not set - cannot solve reCAPTCHA.")
        return False



    # --- Extract sitekey from page ---
    sitekey = None

    # Method 1: data-sitekey attribute
    try:
        el = await page.query_selector("[data-sitekey]")
        if el:
            sitekey = await el.get_attribute("data-sitekey")
    except Exception:
        pass

    # Method 2: parse from reCAPTCHA iframe URL
    if not sitekey:
        for f in page.frames:
            if f.url and "google.com/recaptcha" in f.url and "anchor" in f.url:
                match = re.search(r'[?&]k=([^&]+)', f.url)
                if match:
                    sitekey = match.group(1)
                    break

    # Method 3: search script tags
    if not sitekey:
        try:
            sitekey = await page.evaluate("""() => {
                const scripts = document.querySelectorAll('script');
                for (const s of scripts) {
                    const match = s.textContent.match(/sitekey['":\\s]+['"]([^'"]+)['"]/i);
                    if (match) return match[1];
                }
                return null;
            }""")
        except Exception:
            pass

    if not sitekey:
        print("Could not find reCAPTCHA sitekey on page.")
        return False

    page_url = page.url
    print(f"Found reCAPTCHA sitekey: {sitekey[:12]}...")
    print(f"Submitting to 2Captcha...")

    # --- Submit to 2Captcha ---
    submit_params = urllib.parse.urlencode({
        "key": api_key,
        "method": "userrecaptcha",
        "googlekey": sitekey,
        "pageurl": page_url,
        "json": 1,
    })
    submit_url = f"https://2captcha.com/in.php?{submit_params}"

    try:
        req = urllib.request.Request(submit_url)
        with urllib.request.urlopen(req, timeout=30) as resp:
            result = json.loads(resp.read().decode())
    except Exception as e:
        print(f"2Captcha submit error: {e}")
        return False

    if result.get("status") != 1:
        print(f"2Captcha submit failed: {result}")
        return False

    captcha_id = result["request"]
    print(f"2Captcha task submitted (ID: {captcha_id}). Waiting for solution...")

    # --- Poll for result ---
    await asyncio.sleep(10)  # Initial wait before polling

    poll_params_base = {
        "key": api_key,
        "action": "get",
        "id": captcha_id,
        "json": 1,
    }

    max_polls = 24  # 24 * 5s = 120s max polling
    for attempt in range(max_polls):
        poll_url = f"https://2captcha.com/res.php?{urllib.parse.urlencode(poll_params_base)}"
        try:
            req = urllib.request.Request(poll_url)
            with urllib.request.urlopen(req, timeout=30) as resp:
                result = json.loads(resp.read().decode())
        except Exception as e:
            print(f"  Poll error: {e}")
            await asyncio.sleep(5)
            continue

        if result.get("status") == 1:
            token = result["request"]
            print(f"2Captcha solved! Token received ({len(token)} chars).")
            break
        elif result.get("request") == "CAPCHA_NOT_READY":
            print(f"  Waiting... (poll {attempt + 1}/{max_polls})")
            await asyncio.sleep(5)
        else:
            print(f"2Captcha error: {result}")
            return False
    else:
        print("2Captcha timed out after 2 minutes.")
        return False

    # --- Inject token into page ---
    try:
        inject_result = await page.evaluate("""(token) => {
            const info = { textareas_found: 0, textareas_set: 0, created_hidden: false };

            // Set ALL g-recaptcha-response textareas on the page
            document.querySelectorAll('[id="g-recaptcha-response"], textarea[name="g-recaptcha-response"]').forEach(ta => {
                info.textareas_found++;
                ta.style.display = 'block';
                ta.value = token;
                ta.style.display = 'none';
                if (ta.value.length > 0) info.textareas_set++;
            });

            // Ensure a g-recaptcha-response field exists INSIDE the search form
            const form = document.getElementById('searchForm');
            if (form) {
                let formField = form.querySelector('[name="g-recaptcha-response"]');
                if (!formField || formField.value.length === 0) {
                    // Create/replace a hidden input inside the form
                    if (formField) formField.remove();
                    const hidden = document.createElement('input');
                    hidden.type = 'hidden';
                    hidden.name = 'g-recaptcha-response';
                    hidden.value = token;
                    form.appendChild(hidden);
                    info.created_hidden = true;
                }
            }

            // Override grecaptcha.getResponse() to return our token
            if (typeof grecaptcha !== 'undefined') {
                grecaptcha.getResponse = function() { return token; };
                if (grecaptcha.enterprise) {
                    grecaptcha.enterprise.getResponse = function() { return token; };
                }
            }
            return info;
        }""", token)
        print("Token injected into page.")
    except Exception as e:
        print(f"Warning: Could not set textarea directly: {e}")

    # Trigger the reCAPTCHA callback
    try:
        callback_found = await page.evaluate("""(token) => {
            let found = false;

            // Method 1: Find callback from data-callback attribute on widget
            const widget = document.querySelector('.g-recaptcha[data-callback]');
            if (widget) {
                const cbName = widget.getAttribute('data-callback');
                if (typeof window[cbName] === 'function') {
                    window[cbName](token);
                    found = true;
                }
            }

            // Method 2: Walk ___grecaptcha_cfg to find callback (deeper search)
            if (!found && typeof ___grecaptcha_cfg !== 'undefined') {
                const clients = ___grecaptcha_cfg.clients;
                for (const key in clients) {
                    const client = clients[key];
                    const walk = (obj, depth) => {
                        if (depth > 10 || !obj || found) return;
                        for (const k in obj) {
                            try {
                                if (typeof obj[k] === 'function' && (k === 'callback' || k === 'success-callback')) {
                                    obj[k](token);
                                    found = true;
                                    return;
                                }
                                if (typeof obj[k] === 'object' && obj[k] !== null) {
                                    walk(obj[k], depth + 1);
                                }
                            } catch(e) {}
                        }
                    };
                    walk(client, 0);
                }
            }

            // Method 3: Try grecaptcha internal callback via widget ID
            if (!found && typeof grecaptcha !== 'undefined') {
                try {
                    // Get all widget IDs and try executing their callbacks
                    const widgets = document.querySelectorAll('.g-recaptcha');
                    widgets.forEach((w, i) => {
                        try { grecaptcha.execute(i); } catch(e) {}
                    });
                } catch(e) {}
            }

            return found;
        }""", token)
        print(f"reCAPTCHA callback triggered (found: {callback_found}).")
    except Exception as e:
        print(f"Warning: Callback trigger issue: {e}")

    await asyncio.sleep(1)

    # Verify it worked by checking the anchor frame
    try:
        for f in page.frames:
            if f.url and "google.com/recaptcha" in f.url and "anchor" in f.url:
                is_checked = await f.evaluate(
                    "document.querySelector('#recaptcha-anchor')?.getAttribute('aria-checked')"
                )
                if is_checked == "true":
                    print("reCAPTCHA SOLVED!")
                    return True
                break
    except Exception:
        pass

    # Even if anchor check fails, token injection + getResponse override should work
    print("reCAPTCHA token injected (anchor check inconclusive - proceeding).")
    return True

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
        print(f"Launching browser for Orange County (Headless: {is_headless})...")

        browser = await p.chromium.launch(
            headless=is_headless,
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

            # Wait for reCAPTCHA to load
            await asyncio.sleep(1)

            # 3. Solve reCAPTCHA
            print("Solving reCAPTCHA...")
            if await solve_recaptcha(page):
                print("reCAPTCHA solved successfully.")
            else:
                print("Failed to solve reCAPTCHA. Aborting.")
                # We exit here to avoid clicking 'Search' with an open captcha popup
                sys.exit(1)
            
            # 4. Search - submit form via native submit
            # Note: button id="action" shadows form.action, so we use
            # HTMLFormElement.prototype.submit to bypass client-side validation
            print("Clicking Search...")
            await page.evaluate("""() => {
                const form = document.getElementById('searchForm');
                if (form) {
                    HTMLFormElement.prototype.submit.call(form);
                }
            }""")

            # Wait for navigation
            try:
                await page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            await asyncio.sleep(2)

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
