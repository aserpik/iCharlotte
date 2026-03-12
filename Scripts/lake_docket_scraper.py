import sys
import asyncio
import time
import os
import re
import json
import traceback
from playwright.async_api import async_playwright
from dotenv import load_dotenv
from pypdf import PdfWriter

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
    solved = False
    try:
        for f in page.frames:
            if f.url and "google.com/recaptcha" in f.url and "anchor" in f.url:
                is_checked = await f.evaluate(
                    "document.querySelector('#recaptcha-anchor')?.getAttribute('aria-checked')"
                )
                if is_checked == "true":
                    print("reCAPTCHA SOLVED (anchor verified)!")
                    solved = True
                break
    except Exception:
        pass

    if solved:
        return True

    # Drupal's captcha module reads g-recaptcha-response on POST.
    # The token is already in the textarea. The form will validate server-side.
    # Ensure the response textarea is properly named and filled.
    try:
        await page.evaluate(f"""(token) => {{
            // Ensure the textarea exists and has the token
            let ta = document.querySelector('textarea[name="g-recaptcha-response"]');
            if (!ta) {{
                // Some Drupal captcha setups use a different wrapper
                ta = document.querySelector('#g-recaptcha-response');
            }}
            if (!ta) {{
                // Create it if missing (shouldn't happen but just in case)
                const form = document.querySelector('#ecp-searchform-form');
                if (form) {{
                    ta = document.createElement('textarea');
                    ta.name = 'g-recaptcha-response';
                    ta.id = 'g-recaptcha-response';
                    ta.style.display = 'none';
                    form.appendChild(ta);
                }}
            }}
            if (ta) {{
                ta.value = token;
            }}

            // Also set captcha_response if Drupal uses that field name
            const cr = document.querySelector('input[name="captcha_response"]');
            if (cr) {{ cr.value = token; }}
        }}""", token)
        print("reCAPTCHA token re-confirmed in form fields.")
    except Exception as e:
        print(f"Warning: Could not re-confirm token: {e}")

    print("reCAPTCHA token injected (proceeding with form submit).")
    return True


# CSS to force-expand collapsible sections and clean up for PDF rendering
EXPAND_CSS = """
    .collapse, .panel-collapse, .accordion-collapse, [class*="collapse"] {
        display: block !important;
        height: auto !important;
        overflow: visible !important;
        visibility: visible !important;
        opacity: 1 !important;
    }
    .collapse.in, .collapse.show { display: block !important; }
    body, .main-container { overflow: visible !important; height: auto !important; }
    @media print {
        .collapse, .panel-collapse, [class*="collapse"] {
            display: block !important;
            height: auto !important;
            overflow: visible !important;
            visibility: visible !important;
        }
        * { overflow: visible !important; }
    }
    /* Hide navigation/header elements in PDF */
    .navbar, .breadcrumb, footer, .footer { display: none !important; }
"""


async def render_tab_to_pdf(page, tab_name, tab_selector, temp_dir, index):
    """Click a tab, wait for content, and render to a temporary PDF."""
    try:
        print(f"  Capturing tab: {tab_name}...")

        # Click the tab
        tab = page.locator(tab_selector)
        if await tab.count() == 0:
            print(f"    Tab not found: {tab_name} (selector: {tab_selector})")
            return None

        await tab.click()
        await page.wait_for_load_state("networkidle")
        await asyncio.sleep(1.5)

        # Force-expand any collapsible sections within this tab
        await page.evaluate("""() => {
            document.querySelectorAll('.collapse, .panel-collapse, .accordion-collapse, [class*="collapse"]').forEach(el => {
                el.style.display = 'block';
                el.style.height = 'auto';
                el.style.overflow = 'visible';
                el.style.visibility = 'visible';
                el.style.opacity = '1';
                el.classList.add('in', 'show');
            });
        }""")

        # Render to PDF
        temp_pdf = os.path.join(temp_dir, f"tab_{index}_{tab_name.replace(' ', '_')}.pdf")
        await page.pdf(path=temp_pdf, format="Letter", print_background=True)
        print(f"    Captured: {tab_name}")
        return temp_pdf

    except Exception as e:
        print(f"    Error capturing tab {tab_name}: {e}")
        return None


async def main():
    if len(sys.argv) < 2:
        print("Usage: python lake_docket_scraper.py <case_number> [--headless] [--headful]")
        sys.exit(1)

    case_number = sys.argv[1]
    is_headless = "--headless" in sys.argv or "--headful" not in sys.argv  # Default headless

    search_url = "https://portal.lake.courts.ca.gov/public-portal/?q=node/393"

    # Tab definitions: (display name, CSS selector)
    tabs = [
        ("Case", "a[href*='node/394']:has-text('Case'), .nav-tabs a:has-text('Case'), a.nav-link:has-text('Case')"),
        ("Filings", "a[href*='Filings']:has-text('Filings'), .nav-tabs a:has-text('Filings'), a.nav-link:has-text('Filings')"),
        ("Parties", "a[href*='Parties']:has-text('Parties'), .nav-tabs a:has-text('Parties'), a.nav-link:has-text('Parties')"),
        ("Documents", "a[href*='Documents']:has-text('Documents'), .nav-tabs a:has-text('Documents'), a.nav-link:has-text('Documents')"),
        ("Events", "a[href*='Events']:has-text('Events'), .nav-tabs a:has-text('Events'), a.nav-link:has-text('Events')"),
        ("Case Transfer", "a[href*='Transfer']:has-text('Case Transfer'), .nav-tabs a:has-text('Case Transfer'), a.nav-link:has-text('Case Transfer')"),
    ]

    # Create temp directory for intermediate PDFs
    temp_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_lake_temp_{int(time.time())}")
    os.makedirs(temp_dir, exist_ok=True)

    async with async_playwright() as p:
        print(f"Launching browser for Lake County (Headless: {is_headless})...")
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
            # --- Step 1: Navigate to search page ---
            print(f"Navigating to {search_url}...")
            await page.goto(search_url, wait_until="networkidle")
            await asyncio.sleep(1)

            # --- Step 2: Enter case number ---
            print(f"Entering case number: {case_number}")
            # Journal Technologies form has: Last Name, First Name, Company Name, Case Number
            # as sequential text inputs. Labels contain huge dropdown HTML that breaks
            # text-based selectors. Use row iteration with first-child-text check.
            filled = False

            # Method 1: Iterate form-group rows, match label's first text node
            try:
                rows = await page.query_selector_all(".form-group")
                for row in rows:
                    label = await row.query_selector("label.searchRow")
                    if label:
                        # Get only the label's own first text node, not child element text
                        label_text = await label.evaluate(
                            "el => el.childNodes[0]?.textContent?.trim() || ''"
                        )
                        if label_text.lower() == "case number":
                            inp = await row.query_selector("input.form-control[type='text']")
                            if inp:
                                await inp.fill(case_number)
                                filled = True
                                print(f"  Filled via row label match: '{label_text}'")
                                break
            except Exception as e:
                print(f"  Method 1 (row iteration) failed: {e}")

            # Method 2: Positional — 4th text input in the search form
            if not filled:
                try:
                    inputs = await page.query_selector_all(
                        "#ecp-searchform-form input.form-control[type='text']"
                    )
                    if len(inputs) >= 4:
                        await inputs[3].fill(case_number)
                        filled = True
                        print("  Filled via positional input (4th text field).")
                except Exception as e:
                    print(f"  Method 2 (positional) failed: {e}")

            if not filled:
                print("ERROR: Could not find Case Number input field.")
                await page.screenshot(path="lake_error_screenshot.png")
                sys.exit(1)

            # --- Step 3: Solve reCAPTCHA ---
            print("Solving reCAPTCHA...")
            captcha_solved = False
            for attempt in range(3):
                captcha_solved = await solve_recaptcha(page)
                if captcha_solved:
                    break
                print(f"reCAPTCHA attempt {attempt + 1} failed. Retrying...")
                await asyncio.sleep(2)

            if not captcha_solved:
                print("ERROR: reCAPTCHA could not be solved after 3 attempts.")
                await page.screenshot(path="lake_error_screenshot.png")
                sys.exit(1)

            # --- Step 4: Submit search ---
            print("Submitting search...")
            await page.screenshot(path="lake_debug_before_submit.png")

            pre_url = page.url

            # Try clicking the Search button first
            submit_clicked = False
            for selector in ["#edit-submit", "input[value='Search']", "button:has-text('Search')", "input[type='submit']"]:
                try:
                    el = page.locator(selector)
                    if await el.count() > 0:
                        print(f"  Found submit button: {selector}")
                        await el.first.click(timeout=10000)
                        submit_clicked = True
                        break
                except Exception:
                    continue

            # Wait for navigation/AJAX response
            try:
                await page.wait_for_load_state("networkidle", timeout=15000)
            except Exception:
                pass
            await asyncio.sleep(2)

            # Check if the page actually changed (form might have blocked submit)
            # Note: reCAPTCHA text is in an iframe, not in body.innerText
            page_text = await page.evaluate("document.body?.innerText?.substring(0, 1000) || ''")
            still_on_search = "case number" in page_text.lower() and "last name" in page_text.lower() and "first name" in page_text.lower()

            if still_on_search:
                print("  Form submission appears blocked. Trying JS form.submit() bypass...")
                # Bypass client-side validation by submitting the form directly via JS
                # This sends the g-recaptcha-response textarea to the server for validation
                try:
                    # Remove all submit event listeners by cloning the form, then submit
                    await page.evaluate("""() => {
                        const form = document.getElementById('ecp-searchform-form');
                        if (form) {
                            // Clone removes all JS event handlers
                            const newForm = form.cloneNode(true);
                            form.parentNode.replaceChild(newForm, form);
                            newForm.submit();
                        }
                    }""")
                    print("  JS form.submit() dispatched.")
                    try:
                        await page.wait_for_load_state("networkidle", timeout=30000)
                    except Exception:
                        pass
                    await asyncio.sleep(2)
                except Exception as e:
                    print(f"  JS form submit failed: {e}")

                # Check again
                page_text2 = await page.evaluate("document.body?.innerText?.substring(0, 1000) || ''")
                still_on_search2 = "case number" in page_text2.lower() and "last name" in page_text2.lower() and "first name" in page_text2.lower()
                if still_on_search2:
                    print("  Still on search page after JS submit. Server may have rejected captcha.")
                    await page.screenshot(path="lake_debug_after_jssubmit.png")
                else:
                    print("  JS form submit succeeded — page navigated.")

            await page.screenshot(path="lake_debug_after_submit.png")
            print(f"  Page URL after submit: {page.url}")

            # Debug: check page content for error messages or results
            page_text = await page.evaluate("document.body?.innerText?.substring(0, 500) || ''")
            print(f"  Page content preview: {page_text[:300]}...")

            # --- Step 5: Click into case result ---
            print("Waiting for search results...")

            # Try multiple selectors for case links
            case_link_selectors = [
                f"a:has-text('{case_number}')",
                "table a[href*='node']",
                ".views-field a",
                "td a",
            ]

            found_link = None
            for selector in case_link_selectors:
                try:
                    await page.wait_for_selector(selector, timeout=10000)
                    found_link = selector
                    print(f"  Found results with selector: {selector}")
                    break
                except Exception:
                    continue

            if not found_link:
                content = await page.content()
                if "no cases" in content.lower() or "no result" in content.lower() or "no match" in content.lower():
                    print(f"No results found for case number: {case_number}")
                else:
                    print("Search results did not load in time.")
                    await page.screenshot(path="lake_error_screenshot.png")
                    print("Debug screenshot saved to lake_error_screenshot.png")
                sys.exit(1)

            # Try exact match first, then fall back to first link
            case_links = await page.query_selector_all(found_link)
            if not case_links:
                print("No case links found in search results.")
                sys.exit(1)

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
            await asyncio.sleep(1.5)

            # --- Step 6: Capture all tabs as PDFs ---
            print("Capturing all tabs...")
            await page.emulate_media(media="screen")

            # Inject CSS for expanded collapsible sections
            await page.add_style_tag(content=EXPAND_CSS)

            tab_pdfs = []
            for i, (tab_name, tab_selectors) in enumerate(tabs):
                # Try each selector variant for this tab
                pdf_path = None
                for selector in tab_selectors.split(", "):
                    selector = selector.strip()
                    pdf_path = await render_tab_to_pdf(page, tab_name, selector, temp_dir, i)
                    if pdf_path:
                        break
                if pdf_path:
                    tab_pdfs.append(pdf_path)

            if not tab_pdfs:
                print("ERROR: No tabs could be captured.")
                sys.exit(1)

            # --- Step 7: Merge tab PDFs into single output ---
            output_filename = f"docket_{case_number}_{time.strftime('%Y.%m.%d')}.pdf"
            print(f"Merging {len(tab_pdfs)} tab PDFs into {output_filename}...")

            writer = PdfWriter()
            for pdf_path in tab_pdfs:
                writer.append(pdf_path)
            writer.write(output_filename)
            writer.close()

            print(f"Successfully created {output_filename}")

        except SystemExit:
            raise
        except Exception as e:
            print(f"Error in Lake County scraper: {e}")
            traceback.print_exc()
            try:
                await page.screenshot(path="lake_error_screenshot.png")
                print("Error screenshot saved to lake_error_screenshot.png")
            except Exception:
                pass
            sys.exit(1)
        finally:
            await browser.close()
            # Cleanup temp directory
            try:
                import shutil
                shutil.rmtree(temp_dir, ignore_errors=True)
            except Exception:
                pass


if __name__ == "__main__":
    asyncio.run(main())
