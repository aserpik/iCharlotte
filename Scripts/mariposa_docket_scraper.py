import sys
import asyncio
import time
import os
import re
import json
import traceback
from playwright.async_api import async_playwright
from dotenv import load_dotenv

import urllib.request
import urllib.parse

# Load environment variables from .env file
_env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '.env')
load_dotenv(_env_path, override=True)

STEALTH_JS = """
    Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
    delete navigator.__proto__.webdriver;

    Object.defineProperty(navigator, 'plugins', {
        get: () => {
            const plugins = [
                {name: 'Chrome PDF Plugin', description: 'Portable Document Format', filename: 'internal-pdf-viewer', length: 1, 0: {type: 'application/x-google-chrome-pdf', suffixes: 'pdf', description: 'Portable Document Format'}},
                {name: 'Chrome PDF Viewer', description: '', filename: 'mhjfbmdgcfjbbpaeojofohoefgiehjai', length: 1, 0: {type: 'application/pdf', suffixes: 'pdf', description: ''}},
                {name: 'Native Client', description: '', filename: 'internal-nacl-plugin', length: 2, 0: {type: 'application/x-nacl', suffixes: '', description: 'Native Client Executable'}, 1: {type: 'application/x-pnacl', suffixes: '', description: 'Portable Native Client Executable'}}
            ];
            plugins.item = (i) => plugins[i];
            plugins.namedItem = (name) => plugins.find(p => p.name === name);
            plugins.refresh = () => {};
            return plugins;
        }
    });

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

    Object.defineProperty(navigator, 'languages', {get: () => ['en-US', 'en']});
    Object.defineProperty(navigator, 'language', {get: () => 'en-US'});
    Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
    Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});
    Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});
    Object.defineProperty(navigator, 'connection', {
        get: () => ({effectiveType: '4g', rtt: 50, downlink: 10, saveData: false})
    });

    const originalQuery = window.navigator.permissions.query;
    window.navigator.permissions.query = (parameters) => (
        parameters.name === 'notifications' ?
            Promise.resolve({ state: Notification.permission }) :
            originalQuery(parameters)
    );

    window.chrome = {
        runtime: {connect: () => {}, sendMessage: () => {}, onMessage: {addListener: () => {}, removeListener: () => {}}, onConnect: {addListener: () => {}, removeListener: () => {}}, id: 'mhjfbmdgcfjbbpaeojofohoefgiehjai'},
        loadTimes: () => ({commitLoadTime: Date.now()/1000-5, connectionInfo: 'h2', finishDocumentLoadTime: Date.now()/1000-2, finishLoadTime: Date.now()/1000-1, firstPaintAfterLoadTime: 0, firstPaintTime: Date.now()/1000-3, navigationType: 'Other', npnNegotiatedProtocol: 'h2', requestTime: Date.now()/1000-6, startLoadTime: Date.now()/1000-5, wasAlternateProtocolAvailable: false, wasFetchedViaSpdy: true, wasNpnNegotiated: true}),
        csi: () => ({onloadT: Date.now(), pageT: Date.now()-5000, startE: Date.now()-5000, tran: 15}),
        app: {isInstalled: false, InstallState: {DISABLED: 'disabled', INSTALLED: 'installed', NOT_INSTALLED: 'not_installed'}, RunningState: {CANNOT_RUN: 'cannot_run', READY_TO_RUN: 'ready_to_run', RUNNING: 'running'}}
    };

    const originalContentWindow = Object.getOwnPropertyDescriptor(HTMLIFrameElement.prototype, 'contentWindow');
    Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', {
        get: function() {
            const win = originalContentWindow.get.call(this);
            if (win) { try { Object.defineProperty(win.navigator, 'webdriver', {get: () => undefined}); } catch (e) {} }
            return win;
        }
    });

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

    const oldToString = Function.prototype.toString;
    function newToString() {
        if (this === window.navigator.permissions.query) return 'function query() { [native code] }';
        return oldToString.call(this);
    }
    Function.prototype.toString = newToString;
"""


async def solve_recaptcha(page):
    """Solve reCAPTCHA v2 using 2Captcha token-based API."""

    api_key = os.environ.get("TWOCAPTCHA_API_KEY")
    if not api_key:
        print("TWOCAPTCHA_API_KEY not set - cannot solve reCAPTCHA.")
        return False
    if len(api_key) != 32:
        print(f"WARNING: TWOCAPTCHA_API_KEY has unexpected length {len(api_key)} (expected 32)")

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
    token = None
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
        await page.evaluate(f"""(token) => {{
            // Set ALL g-recaptcha-response textareas on the page
            document.querySelectorAll('[id="g-recaptcha-response"], textarea[name="g-recaptcha-response"]').forEach(ta => {{
                ta.style.display = 'block';
                ta.value = token;
                ta.style.display = 'none';
            }});

            // Override grecaptcha.getResponse() to return our token
            if (typeof grecaptcha !== 'undefined') {{
                const origGetResponse = grecaptcha.getResponse;
                grecaptcha.getResponse = function() {{ return token; }};
                if (grecaptcha.enterprise) {{
                    grecaptcha.enterprise.getResponse = function() {{ return token; }};
                }}
            }}
        }}""", token)
        print("Token injected into page.")
    except Exception as e:
        print(f"Warning: Could not set textarea directly: {e}")

    # Trigger the reCAPTCHA callback
    try:
        callback_found = await page.evaluate(f"""(token) => {{
            let found = false;

            // Method 1: Find callback from data-callback attribute on widget
            const widget = document.querySelector('.g-recaptcha[data-callback]');
            if (widget) {{
                const cbName = widget.getAttribute('data-callback');
                if (typeof window[cbName] === 'function') {{
                    window[cbName](token);
                    found = true;
                }}
            }}

            // Method 2: Walk ___grecaptcha_cfg to find callback
            if (!found && typeof ___grecaptcha_cfg !== 'undefined') {{
                const clients = ___grecaptcha_cfg.clients;
                for (const key in clients) {{
                    const client = clients[key];
                    const walk = (obj, depth) => {{
                        if (depth > 6 || !obj || found) return;
                        for (const k in obj) {{
                            try {{
                                if (k === 'callback' && typeof obj[k] === 'function') {{
                                    obj[k](token);
                                    found = true;
                                    return;
                                }}
                                if (typeof obj[k] === 'object' && obj[k] !== null) {{
                                    walk(obj[k], depth + 1);
                                }}
                            }} catch(e) {{}}
                        }}
                    }};
                    walk(client, 0);
                }}
            }}

            return found;
        }}""", token)
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
    if len(sys.argv) < 2:
        print("Usage: python mariposa_docket_scraper.py <case_number> [--headless] [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv  # Default headless

    url = "https://portal-camariposa.tylertech.cloud/Portal/Home/Dashboard/26"

    async with async_playwright() as p:
        print(f"Launching browser for Mariposa County Superior Court (Headless: {is_headless})...")
        browser = await p.chromium.launch(
            headless=is_headless,
            channel="chrome",
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
                "--disable-setuid-sandbox"
            ]
        )
        context = await browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
            accept_downloads=True
        )

        # Inject stealth scripts before any page loads
        await context.add_init_script(STEALTH_JS)

        page = await context.new_page()

        try:
            # --- Step 1: Navigate to portal ---
            print(f"Navigating to {url}...")
            await page.goto(url, wait_until="networkidle", timeout=60000)
            await asyncio.sleep(0.5)

            # --- Step 2: Select "Case Number" search type ---
            print("Selecting 'Case Number' search type...")
            await page.select_option("select#cboHSSearchBy", value="CaseNumber")

            # --- Step 3: Input case number ---
            print(f"Inputting case number: {case_number}")
            await page.fill("input#SearchCriteria_SearchValue", case_number)

            # --- Step 4: Input date range ---
            print("Inputting 'Date From' (01/01/2025)...")
            await page.fill("input#SearchCriteria_DateFrom", "01/01/2025")

            print("Inputting 'Date To' (12/01/2026)...")
            await page.fill("input#SearchCriteria_DateTo", "12/01/2026")

            # --- Step 5: Solve reCAPTCHA ---
            captcha_solved = await solve_recaptcha(page)
            if not captcha_solved:
                print("ERROR: reCAPTCHA could not be solved.")
                await page.screenshot(path=f"error_mariposa_{case_number}.png")
                sys.exit(1)

            # --- Step 6: Click Submit ---
            print("Clicking Submit...")
            try:
                await page.click("input#btnHSSubmit", force=True)
            except Exception:
                # Fallback: press Enter on the search field
                await page.focus("input#SearchCriteria_SearchValue")
                await page.keyboard.press("Enter")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(1)

            # --- Step 7: Click the case link ---
            print("Searching for case link...")
            try:
                await page.wait_for_selector("a.caseLink", state="visible", timeout=30000)
            except Exception:
                content = await page.content()
                if "no cases" in content.lower() or "no result" in content.lower():
                    print(f"No results found for case number: {case_number}")
                else:
                    print("Search results did not load in time.")
                    await page.screenshot(path=f"error_mariposa_{case_number}.png")
                sys.exit(1)

            case_links = await page.query_selector_all("a.caseLink")
            if not case_links:
                print("No case links found in search results.")
                sys.exit(1)

            # Try exact match first
            clicked = False
            for link in case_links:
                link_text = (await link.inner_text()).strip()
                if link_text.upper() == case_number.upper():
                    print(f"Found exact match: {link_text}")
                    await link.click()
                    clicked = True
                    break

            if not clicked:
                link_text = (await case_links[0].inner_text()).strip()
                print(f"Clicking first case link: {link_text}")
                await case_links[0].click()

            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(0.5)

            # --- Step 8: Generate PDF ---
            print("Generating PDF...")
            date_str = time.strftime("%Y.%m.%d")
            output_filename = f"docket_{case_number}_{date_str}.pdf"

            await page.pdf(path=output_filename, format="Letter", print_background=True)
            print(f"Successfully created {output_filename}")

        except SystemExit:
            raise
        except Exception as e:
            print(f"An error occurred: {e}")
            traceback.print_exc()
            try:
                if not page.is_closed():
                    await page.screenshot(path=f"error_mariposa_{case_number}.png")
                    print(f"Error screenshot saved to error_mariposa_{case_number}.png")
            except Exception:
                pass
            sys.exit(1)
        finally:
            await browser.close()

if __name__ == "__main__":
    asyncio.run(main())
