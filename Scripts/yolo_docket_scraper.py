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
                    // The callback is typically at client[key][key].callback
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
        print("Usage: python yolo_docket_scraper.py <case_number> [--headless] [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv  # Default headless

    portal_url = "https://portal-cayolo.tylertech.cloud/Portal/Home/Dashboard/29"

    async with async_playwright() as p:
        print(f"Launching browser for Yolo County (Headless: {is_headless})...")
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
            print(f"Navigating to {portal_url}...")
            await page.goto(portal_url, wait_until="networkidle")
            await asyncio.sleep(0.5)

            # --- Step 2: Enter case number ---
            print(f"Entering case number: {case_number}")
            await page.wait_for_selector("#caseCriteria_SearchCriteria", state="visible", timeout=15000)
            await page.fill("#caseCriteria_SearchCriteria", case_number)

            # --- Step 3: Solve reCAPTCHA ---
            captcha_solved = await solve_recaptcha(page)
            if not captcha_solved:
                print("ERROR: reCAPTCHA could not be solved.")
                await page.screenshot(path="error_screenshot.png")
                sys.exit(1)

            # --- Step 4: Click Submit ---
            print("Clicking Submit...")
            await page.click("#btnSSSubmit")
            await page.wait_for_load_state("networkidle")
            await asyncio.sleep(0.5)

            # --- Step 5: Click case link ---
            print("Waiting for search results...")
            try:
                await page.wait_for_selector("a.caseLink", state="visible", timeout=15000)
            except Exception:
                content = await page.content()
                if "no cases" in content.lower() or "no result" in content.lower():
                    print(f"No results found for case number: {case_number}")
                else:
                    print("Search results did not load in time.")
                    await page.screenshot(path="error_screenshot.png")
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

            # --- Step 6: Click Print button to show print options ---
            print("Looking for Print button...")
            # Override window.print before clicking anything to prevent browser dialog
            await page.evaluate("window.__originalPrint = window.print; window.print = () => {};")

            # The page may have different print button variants:
            # 1. toggleElementVisible('PrintMask') button (shows print section options)
            # 2. Direct window.print() button
            try:
                # First try the toggleElementVisible button (user-described flow)
                toggle_btn = page.locator("button[onclick*='toggleElementVisible']")
                if await toggle_btn.count() > 0:
                    print("Found Print toggle button, clicking...")
                    await toggle_btn.first.click(force=True)
                    await asyncio.sleep(0.5)

                    # Now click SectionPrintButton
                    section_btn = page.locator("#SectionPrintButton")
                    if await section_btn.count() > 0 and await section_btn.is_visible():
                        print("Clicking Section Print button...")
                        await section_btn.click(force=True)
                        await asyncio.sleep(1)
                else:
                    # Fallback: click any visible Print button
                    print("No toggle button found. Clicking Print button directly...")
                    print_btns = page.locator("button:has-text('Print')")
                    count = await print_btns.count()
                    for i in range(count):
                        btn = print_btns.nth(i)
                        if await btn.is_visible():
                            await btn.click(force=True)
                            await asyncio.sleep(1)
                            break
            except Exception as e:
                print(f"Print button interaction: {e}")

            # --- Step 7: Save as PDF ---
            output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Saving page as PDF: {output_filename}")
            await page.pdf(path=output_filename, format="Letter", print_background=True)
            print(f"Successfully saved: {output_filename}")

        except SystemExit:
            raise
        except Exception as e:
            print(f"Error in Yolo County scraper: {e}")
            traceback.print_exc()
            try:
                await page.screenshot(path="error_screenshot.png")
                print("Error screenshot saved to error_screenshot.png")
            except Exception:
                pass
            sys.exit(1)
        finally:
            await browser.close()
            # Cleanup temp files
            for f in ["error_screenshot.png"]:
                if os.path.exists(f):
                    try:
                        os.remove(f)
                    except Exception:
                        pass

if __name__ == "__main__":
    asyncio.run(main())
