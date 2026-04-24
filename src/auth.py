"""Automated Graph Explorer token acquisition via Playwright.

Uses browser automation to navigate the Graph Explorer -> Microsoft Login -> Duo SSO
flow, extract the access token, and return it. A persistent browser context is used
to retain cookies between runs so repeat auths can skip the login steps.

The Microsoft login flow opens in a **popup window** from Graph Explorer. All login
and Duo SSO interactions happen in the popup; once auth completes the popup closes
and the main Graph Explorer page becomes authenticated.
"""

import json
import logging
import subprocess
import sys
from pathlib import Path
from urllib.parse import urlparse

logger = logging.getLogger(__name__)

# Persistent browser profile directory
BROWSER_DATA_DIR = Path.home() / ".claude-graph-auth"

# Graph Explorer URL
GRAPH_EXPLORER_URL = "https://developer.microsoft.com/en-us/graph/graph-explorer"

# The Duo approval step requires user interaction (phone push, browser confirm, etc.)
# so it gets a generous timeout independent of the general `timeout` parameter.
DUO_APPROVAL_TIMEOUT = 120  # seconds


def _url_has_domain(url: str, domain: str) -> bool:
    """Check if a URL's hostname matches or is a subdomain of the given domain."""
    try:
        hostname = urlparse(url).hostname or ""
        return hostname == domain or hostname.endswith("." + domain)
    except Exception:
        return False


def _ensure_playwright() -> str | None:
    """Ensure Playwright and its Chromium browser are installed.

    Installs into the current Python environment (venv) if missing.
    Returns None on success, or an error message string on failure.
    """
    try:
        import playwright  # noqa: F401
    except ImportError:
        print("Installing playwright...", file=sys.stderr)
        result = subprocess.run(
            [sys.executable, "-m", "pip", "install", "playwright"],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            return f"Failed to pip install playwright: {result.stderr.strip()}"

    print("Ensuring Chromium browser is installed...", file=sys.stderr)
    result = subprocess.run(
        [sys.executable, "-m", "playwright", "install", "chromium"],
        capture_output=True, text=True,
    )
    if result.returncode != 0:
        return f"Failed to install Chromium browser: {result.stderr.strip()}"

    return None


def authenticate(email: str = None, timeout: int = 10) -> dict:
    """Automate Graph Explorer auth flow and extract access token.

    Opens a headed Chromium browser with a persistent profile, navigates to
    Graph Explorer, goes through Microsoft login + Duo SSO (in a popup window),
    and extracts the access token.

    Automatically installs Playwright and Chromium into the current venv if
    they are not already present.

    Args:
        email: Microsoft account email address. If None, the user must type it.
        timeout: Seconds to wait for non-interactive steps (default 10).
            The Duo approval step uses DUO_APPROVAL_TIMEOUT (120s) or this
            value, whichever is larger.

    Returns:
        dict with keys:
            - status: "success" or "error"
            - token: The access token (only on success)
            - step: Which step failed (only on error)
            - message: Error description (only on error)
    """
    install_err = _ensure_playwright()
    if install_err:
        return {"status": "error", "step": "install", "message": install_err}

    from playwright.sync_api import sync_playwright

    BROWSER_DATA_DIR.mkdir(parents=True, exist_ok=True)
    _clean_crash_state()
    _seed_browser_preferences()

    with sync_playwright() as p:
        context = None
        try:
            context = p.chromium.launch_persistent_context(
                user_data_dir=str(BROWSER_DATA_DIR),
                headless=False,
                args=[
                    "--disable-blink-features=WebAuthentication",
                    "--disable-features=PrivateNetworkAccessPermissionPrompt",
                    "--hide-crash-restore-bubble",
                    "--disable-session-crashed-bubble",
                    "--no-first-run",
                    "--disable-infobars",
                ],
                ignore_default_args=["--enable-automation"],
            )
            page = context.pages[0] if context.pages else context.new_page()

            # ── Step 1: Navigate to Graph Explorer ──────────────────────
            print("Opening Graph Explorer...", file=sys.stderr)
            page.goto(GRAPH_EXPLORER_URL, wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(3000)  # let React app settle

            # ── Step 2: Check if already authenticated ──────────────────
            token = _try_extract_token(page)
            if token:
                print("Already authenticated — token extracted.", file=sys.stderr)
                return {"status": "success", "token": token}

            # ── Step 3: Click sign-in and capture popup ─────────────────
            print("Clicking sign-in (opens popup)...", file=sys.stderr)
            popup = _click_sign_in_and_get_popup(page, context)
            if popup is None:
                _save_debug_screenshot(page, "no_popup")
                return _error("sign_in", "Could not open login popup from Graph Explorer")

            print(f"Login popup opened: {popup.url}", file=sys.stderr)

            # Inject a script to disable WebAuthn in the popup so
            # Duo's "Use your security key" page can't trigger a blocking
            # browser dialog that swallows all clicks.
            popup.add_init_script("""
                if (navigator.credentials) {
                    navigator.credentials.get = () =>
                        Promise.reject(new DOMException('Automated', 'NotAllowedError'));
                    navigator.credentials.create = () =>
                        Promise.reject(new DOMException('Automated', 'NotAllowedError'));
                }
            """)

            # ── Step 4: Microsoft login (in popup) ──────────────────────
            print("Handling Microsoft login in popup...", file=sys.stderr)
            result = _handle_microsoft_login(popup, email)
            if result is not None:
                return result

            # ── Step 5: Duo SSO (in popup) ──────────────────────────────
            print("Handling Duo SSO in popup...", file=sys.stderr)
            result = _handle_duo_sso(popup)
            if result is not None:
                return result

            # ── Step 6: Wait for popup to close (auth complete) ─────────
            # The Duo approval step requires user interaction (phone push, etc.)
            # so use the dedicated DUO_APPROVAL_TIMEOUT as the minimum.
            popup_timeout = max(timeout, DUO_APPROVAL_TIMEOUT)
            print(
                f"Waiting for auth to complete (up to {popup_timeout}s)...",
                file=sys.stderr,
            )
            result = _wait_for_popup_close(popup, page, popup_timeout)
            if result is not None:
                return result

            # ── Step 7: Extract access token from main page ─────────────
            print("Extracting access token...", file=sys.stderr)
            import time
            time.sleep(3)  # let Graph Explorer process the auth callback

            # Re-acquire the page from context — the original reference may be stale
            page = None
            for pg in context.pages:
                try:
                    if "graph-explorer" in pg.url or _url_has_domain(pg.url, "developer.microsoft.com"):
                        page = pg
                        break
                except Exception:
                    continue
            if page is None:
                # No Graph Explorer page found — open a fresh one
                page = context.new_page()
                page.goto(GRAPH_EXPLORER_URL, wait_until="domcontentloaded", timeout=30000)
                page.wait_for_timeout(5000)
            else:
                # Reload to ensure the auth state is reflected
                try:
                    page.reload(wait_until="domcontentloaded", timeout=15000)
                    page.wait_for_timeout(3000)
                except Exception:
                    page.goto(GRAPH_EXPLORER_URL, wait_until="domcontentloaded", timeout=30000)
                    page.wait_for_timeout(5000)

            token = _extract_token(page)
            if not token:
                _save_debug_screenshot(page, "token_extraction_failed")
                return _error("token_extract", "Could not extract access token from Graph Explorer")

            print("Authentication successful!", file=sys.stderr)
            return {"status": "success", "token": token}

        except Exception as e:
            if context:
                try:
                    _save_debug_screenshot(
                        context.pages[0] if context.pages else None,
                        "unexpected_error",
                    )
                except Exception:
                    pass
            return _error("unexpected", str(e))
        finally:
            if context:
                context.close()


# ─── Internal helpers ────────────────────────────────────────────────────────


def _clean_crash_state() -> None:
    """Remove crash indicators so Chromium doesn't show 'Restore pages?' dialog."""
    for name in ["SingletonLock", "SingletonSocket", "SingletonCookie"]:
        lock = BROWSER_DATA_DIR / name
        if lock.exists():
            try:
                lock.unlink()
            except Exception:
                pass
    # Clear the exit_type in Preferences to prevent "didn't shut down correctly"
    prefs_file = BROWSER_DATA_DIR / "Default" / "Preferences"
    if prefs_file.exists():
        try:
            prefs = json.loads(prefs_file.read_text())
            if "profile" in prefs:
                prefs["profile"]["exit_type"] = "Normal"
                prefs["profile"]["exited_cleanly"] = True
                prefs_file.write_text(json.dumps(prefs, indent=2))
        except Exception:
            pass


def _seed_browser_preferences() -> None:
    """Pre-seed browser preferences to auto-grant private network access.

    Duo Desktop communicates with the Duo website via a local connection.
    Chromium prompts the user to allow this ("Access other apps and services
    on this device"). We pre-grant this permission for duosecurity.com origins
    in the browser's Preferences file so the prompt never appears.
    """
    prefs_dir = BROWSER_DATA_DIR / "Default"
    prefs_dir.mkdir(parents=True, exist_ok=True)
    prefs_file = prefs_dir / "Preferences"

    # Load existing preferences or start fresh
    prefs = {}
    if prefs_file.exists():
        try:
            prefs = json.loads(prefs_file.read_text())
        except Exception:
            prefs = {}

    # Navigate/create the nested dict for content settings
    profile = prefs.setdefault("profile", {})
    content_settings = profile.setdefault("content_settings", {})
    exceptions = content_settings.setdefault("exceptions", {})

    # Grant private network access for duosecurity.com
    # Chromium uses "private-network-access" as the content setting key.
    # The origin pattern format is "origin,embeddingOrigin" (comma-separated).
    # Setting 1 = ALLOW, 2 = BLOCK.
    pna = exceptions.setdefault("private-network-access", {})
    for pattern in [
        "https://[*.]duosecurity.com:443,*",
        "https://[*.]duosecurity.com,*",
    ]:
        if pattern not in pna:
            pna[pattern] = {
                "expiration": "0",
                "last_modified": "0",
                "model": 0,
                "setting": 1,
            }

    try:
        prefs_file.write_text(json.dumps(prefs, indent=2))
    except Exception as e:
        logger.warning("Could not write browser preferences: %s", e)


def _error(step: str, message: str) -> dict:
    logger.error("Auth failed at step '%s': %s", step, message)
    return {"status": "error", "step": step, "message": message}


def _save_debug_screenshot(page, name: str) -> None:
    if page is None:
        return
    try:
        path = BROWSER_DATA_DIR / f"debug_{name}.png"
        page.screenshot(path=str(path))
        print(f"Debug screenshot saved to {path}", file=sys.stderr)
    except Exception:
        pass


def _click_sign_in_and_get_popup(page, context):
    """Click the sign-in button and return the popup Page that opens.

    Graph Explorer opens a popup window for the Microsoft login flow.
    Returns the popup Page, or None if no popup appeared.
    """
    try:
        # Set up popup listener before clicking
        with page.expect_popup(timeout=15000) as popup_info:
            # Try various sign-in button selectors
            for selector in [
                page.get_by_role("button", name="Sign in"),
                page.get_by_text("Sign in to Graph Explorer"),
                page.get_by_text("Sign in"),
                page.locator("[aria-label*='sign in' i]"),
                page.locator("button:has-text('Sign in')"),
            ]:
                if selector.count() > 0:
                    selector.first.click()
                    break
        popup = popup_info.value
        # The popup starts as about:blank then redirects to login.microsoftonline.com.
        # Wait for it to navigate to the actual login page.
        popup.wait_for_load_state("domcontentloaded", timeout=15000)
        # Wait until the URL is no longer about:blank
        if popup.url == "about:blank":
            popup.wait_for_url("**/login.microsoftonline.com/**", timeout=15000)
            popup.wait_for_load_state("domcontentloaded", timeout=15000)
        return popup
    except Exception as e:
        logger.warning("Error getting login popup: %s", e)
        return None


def _handle_microsoft_login(popup, email: str | None) -> dict | None:
    """Handle the Microsoft login flow inside the popup window.

    Returns None on success (moved past login), or an error dict on failure.
    """
    try:
        # If cached credentials cause the popup to close instantly, that
        # means auth completed successfully — treat it as a fast-path.
        if popup.is_closed():
            print("Popup already closed (cached auth) — skipping login.", file=sys.stderr)
            return None

        popup.wait_for_timeout(2000)

        if popup.is_closed():
            print("Popup closed during wait (cached auth) — skipping login.", file=sys.stderr)
            return None

        # Check for account picker (returning user with cached account)
        account_picker = popup.locator("[data-test-id*='@']")
        if account_picker.count() > 0:
            if email:
                # Try clicking the specific account
                specific = popup.locator(f"[data-test-id='{email}']")
                if specific.count() > 0:
                    specific.first.click()
                    popup.wait_for_timeout(3000)
                    return None
            # Click the first account
            account_picker.first.click()
            popup.wait_for_timeout(3000)
            return None

        # Check for email input field (fresh login)
        email_input = popup.locator("input[type='email'], input[name='loginfmt']")
        if email_input.count() > 0:
            if not email:
                print(
                    "Email input detected but no email configured. "
                    "Please type your email in the popup.",
                    file=sys.stderr,
                )
                # Wait for user to manually proceed past login
                try:
                    popup.wait_for_url("**/*duo*/**", timeout=60000)
                except Exception:
                    pass
                return None

            email_input.first.fill(email)
            popup.wait_for_timeout(500)

            # Click Next
            next_btn = popup.locator("input[type='submit'][value='Next'], input[type='submit']")
            if next_btn.count() == 0:
                next_btn = popup.get_by_role("button", name="Next")
            if next_btn.count() > 0:
                next_btn.first.click()
                popup.wait_for_timeout(3000)
            return None

        # May have skipped straight to Duo (cached credentials)
        return None

    except Exception as e:
        # A closed popup during login means auth completed via cached
        # credentials — this is success, not an error.
        if popup.is_closed():
            print("Popup closed (cached auth) — treating as success.", file=sys.stderr)
            return None
        _save_debug_screenshot(popup, "microsoft_login_failed")
        return _error("microsoft_login", str(e))


def _handle_duo_sso(popup) -> dict | None:
    """Handle Duo SSO inside the popup window.

    Tries to initiate the auth (push, system browser, etc.).
    Returns None on success, or an error dict on failure.
    """
    try:
        # Popup may have already closed (fast cached auth) — that's success.
        if popup.is_closed():
            print("Popup already closed before Duo — skipping.", file=sys.stderr)
            return None

        # Wait for the Duo page to fully render (it uses client-side JS)
        try:
            popup.wait_for_load_state("networkidle", timeout=10000)
        except Exception:
            if popup.is_closed():
                return None
            popup.wait_for_timeout(3000)

        # The popup may still be on the Microsoft login page transitioning to Duo.
        # Wait briefly for a possible navigation to duosecurity.com.
        try:
            popup.wait_for_url("**/*duo*/**", timeout=10000)
        except Exception:
            pass  # might already be there or might not go to Duo at all

        url = popup.url
        is_duo = _url_has_domain(url, "duosecurity.com") or _url_has_domain(url, "duo.com")
        if not is_duo and popup.locator("#duo_iframe").count() > 0:
            is_duo = True
        if not is_duo:
            logger.info("Not on Duo page (url: %s) — skipping", url)
            return None

        print(f"On Duo page: {url[:80]}", file=sys.stderr)

        # The Duo flow has multiple screens. Loop through them,
        # clicking the right button on each screen until the popup closes
        # or we run out of known buttons.
        _navigate_duo_screens(popup)
        return None

    except Exception as e:
        logger.warning("Non-fatal error handling Duo: %s", e)
        return None


def _navigate_duo_screens(popup) -> None:
    """Click through the sequence of Duo screens until auth completes.

    The Duo flow can show several pages in sequence:
    1. "Log in" button (SSO page with remembered email)
    2. "Skip for now" (Duo Desktop local network permission)
    3. "Use your security key" → fails → "Show other options"
    4. Auth method list → select "Duo Desktop" or "Duo Push"
    5. "Continue in browser" button
    6. Popup closes (auth complete)

    This function loops up to 10 rounds, clicking whatever action
    button is visible on each screen.
    """
    # Ordered list of button texts to look for on each screen.
    # We try them all each round — the first visible one gets clicked.
    action_buttons = [
        # Initial login
        "Log in",
        # Skip Duo Desktop permission
        "Skip for now",
        # After security key fails
        "Show other options",
        # Method selection (handled specially below)
        # Continue in browser
        "Continue in browser",
        # Other common buttons
        "Send Me a Push",
        "Next",
        "Yes, this is my device",
        "Trust this browser",
    ]

    for round_num in range(10):
        if popup.is_closed():
            return

        try:
            popup.wait_for_load_state("domcontentloaded", timeout=3000)
        except Exception:
            pass
        popup.wait_for_timeout(1000)

        if popup.is_closed():
            return

        # Check if Duo shows "Success" — auth is done, just wait for redirect
        try:
            if popup.get_by_text("Success").count() > 0:
                print("Duo auth succeeded! Waiting for redirect...", file=sys.stderr)
                try:
                    popup.wait_for_event("close", timeout=10000)
                except Exception:
                    pass
                return
        except Exception:
            pass

        clicked = False

        # Check for auth method selection page (after "Show other options")
        # This page has a list of methods — select one
        try:
            for method_text in ["Duo Desktop", "Duo Mobile", "Duo Push", "Verified Duo Push"]:
                opt = popup.get_by_text(method_text)
                if opt.count() > 0:
                    # Only click if we're actually on a method selection page
                    # (i.e., "Other options" or similar is also visible, indicating a list)
                    if (popup.get_by_text("Secured by Duo").count() > 0
                            and popup.get_by_text("Other options").count() > 0):
                        print(f"Selecting auth method '{method_text}'...", file=sys.stderr)
                        opt.first.click()
                        clicked = True
                        popup.wait_for_timeout(2000)
                        break
        except Exception:
            pass

        if not clicked:
            # Try each action button
            for btn_text in action_buttons:
                try:
                    locator = popup.get_by_text(btn_text, exact=True)
                    if locator.count() > 0 and locator.first.is_visible():
                        print(f"[Duo step {round_num + 1}] Clicking '{btn_text}'...", file=sys.stderr)
                        locator.first.click()
                        clicked = True
                        popup.wait_for_timeout(2000)
                        break
                except Exception:
                    continue

        if not clicked:
            # Nothing to click — either waiting for user or unknown page
            _save_debug_screenshot(popup, f"duo_round_{round_num}")
            print(f"[Duo step {round_num + 1}] No known button found, waiting...", file=sys.stderr)


def _wait_for_popup_close(popup, main_page, timeout: int) -> dict | None:
    """Wait for the login popup to close, indicating auth is complete.

    The popup closes automatically once the Microsoft/Duo auth flow finishes
    and the token is passed back to Graph Explorer.

    Returns None on success, or an error dict on timeout.
    """
    import time

    deadline = time.monotonic() + timeout
    poll_interval = 2

    while time.monotonic() < deadline:
        try:
            if popup.is_closed():
                print("Login popup closed — auth complete.", file=sys.stderr)
                return None
        except Exception:
            return None

        # While waiting, click through any post-auth prompts in the popup
        # (e.g., Microsoft "Stay signed in?" page)
        try:
            for btn_text in ["Yes", "No, show this again", "Accept"]:
                locator = popup.get_by_text(btn_text, exact=True)
                if locator.count() > 0 and locator.first.is_visible():
                    # Only click "Yes" for "Stay signed in?"
                    if btn_text == "Yes":
                        print("Clicking 'Yes' on 'Stay signed in?' prompt...", file=sys.stderr)
                    locator.first.click()
                    popup.wait_for_timeout(1000)
                    break
        except Exception:
            pass

        try:
            main_page.wait_for_timeout(poll_interval * 1000)
        except Exception:
            time.sleep(poll_interval)

    # Timeout — take screenshots of both windows for debugging
    _save_debug_screenshot(main_page, "auth_timeout_main")
    try:
        if not popup.is_closed():
            _save_debug_screenshot(popup, "auth_timeout_popup")
    except Exception:
        pass

    return _error(
        "auth_timeout",
        f"Authentication did not complete within {timeout} seconds. "
        "Check the browser popup and try again.",
    )


def _try_extract_token(page) -> str | None:
    """Try to extract a token if the user is already authenticated.

    Clicks the Access Token tab and reads the token text.
    """
    try:
        # Look for the Access Token tab button
        token_tab = page.get_by_role("tab", name="Access token")
        if token_tab.count() == 0:
            token_tab = page.locator("button:has-text('Access token')")
        if token_tab.count() == 0:
            return None

        token_tab.first.click()
        page.wait_for_timeout(1000)

        return _find_jwt_on_page(page)
    except Exception:
        pass
    return None


def _extract_token(page) -> str | None:
    """Extract the access token from Graph Explorer after auth completes."""
    try:
        # Try the quick path first
        token = _try_extract_token(page)
        if token:
            return token

        # Give the page more time
        page.wait_for_timeout(2000)

        # Try clicking the Access token tab with various selectors
        for selector in [
            page.get_by_role("tab", name="Access token"),
            page.get_by_text("Access token").first,
            page.locator("button:has-text('Access token')"),
        ]:
            try:
                if selector.count() > 0:
                    selector.first.click()
                    page.wait_for_timeout(1500)
                    break
            except Exception:
                continue

        token = _find_jwt_on_page(page)
        if token:
            return token

        # Last resort: JS evaluation to find any JWT in the DOM
        try:
            token = page.evaluate("""() => {
                const walk = document.createTreeWalker(
                    document.body, NodeFilter.SHOW_TEXT, null);
                while (walk.nextNode()) {
                    const text = walk.currentNode.textContent.trim();
                    if (text.startsWith('eyJ') && text.length > 100) {
                        return text;
                    }
                }
                return null;
            }""")
            if token:
                return token.strip()
        except Exception:
            pass

    except Exception as e:
        logger.warning("Error extracting token: %s", e)

    return None


def _find_jwt_on_page(page) -> str | None:
    """Search common containers for a JWT token string."""
    for selector in [
        "[class*='access-token']",
        "[class*='token']",
        "[class*='Token']",
        "textarea",
        "pre",
        "[role='textbox']",
        ".monaco-editor .view-lines",
    ]:
        els = page.locator(selector)
        for i in range(min(els.count(), 5)):
            try:
                text = els.nth(i).inner_text()
                if text and text.strip().startswith("eyJ"):
                    return text.strip()
            except Exception:
                continue
    return None
