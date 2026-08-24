#!/usr/bin/env python3
"""
sso_login.py — one-click Microsoft sign-in for the Jira timesheet web UI.

The user clicks "Sign in with Microsoft" and nothing else. Everything the old
flow asked them to do by hand (open Jira, sign in, F12, copy JSESSIONID, paste
it back) happens here instead:

  1. If a browser on this machine already holds a live Jira session, we borrow
     those cookies and log in instantly.
  2. Otherwise a headless browser walks the SSO redirect on our own profile —
     nothing appears on screen — and the moment Jira hands out a session that
     works, we take it and the app is logged in.

The browser we drive keeps its own profile folder next to this file, so the
Microsoft session persists between runs. Nothing here can type a password,
though: if Microsoft actually prompts, we stop and ask the user to sign in
once in their own browser.

Optional dependencies (both degrade gracefully):
    pip install selenium        # drives the sign-in window  (the main path)
    pip install browser_cookie3 # reuse an already-signed-in browser (shortcut)
"""

from __future__ import annotations

import os
import threading
import time
from typing import Callable, Optional
from urllib.parse import urlparse

import jira_credentials
import jira_logging_utility as core

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, ".sso-browser-profile")

POLL_SECONDS = 1.0
SILENT_TIMEOUT = 45            # a silent SSO round trip, not a human typing
WINDOW_TIMEOUT = 300           # a human signing in, when they ask for a window
PAGE_LOAD_TIMEOUT = 25         # never block forever inside driver.get()
PROBE_TIMEOUT = 8              # per cookie-probe request

# Nothing can be clicked in the background, so when Microsoft actually wants a
# password we stop and say so rather than hanging.
_NOT_SIGNED_IN = (
    "Couldn't sign in silently — it stopped at {where}. Either sign in to {url} "
    "with Microsoft in your own browser and click again, or open a sign-in "
    "window here."
)

_LOGIN_HOSTS = ("login.microsoftonline.com", "login.live.com", "login.windows.net",
                "adfs", "sts.", "okta", "/login", "signin", "sign-in", "auth")

# These hosts use the stock Atlassian username/password form — there is no
# Microsoft SSO to walk, so we sign in with the saved credentials instead.
PASSWORD_ONLY_HOSTS = ("tracking.i2cinc.com",)

_SAVED_REJECTED = (
    "The saved sign-in for {where} was rejected — the password or token has "
    "probably changed. Enter it again above and press Sign in; it will be "
    "saved over the old one."
)

_NEED_CREDENTIALS = (
    "Nothing saved yet for {where}. Enter your username and password (or a "
    "Personal Access Token) above and press Sign in once — after that this "
    "button signs you in on its own, with no typing and no browser."
)

# The stock Atlassian login form — or, on the anonymous dashboard, the plain
# "Log in" link that stands in for it. Either means: no SSO is coming.
_JS_JIRA_LOGIN_FORM = """
return !!(document.getElementById('login-form-username') ||
          document.getElementById('username-field') ||
          document.querySelector('form#login-form input[type=password]') ||
          document.querySelector("a[href*='login.jsp']"));
"""
_SSO_GRACE = 6      # seconds to let a redirect happen before judging the page

# One profile means one browser at a time: the sign-in and the attendance fetch
# take turns rather than colliding over a locked user-data-dir.
BROWSER_LOCK = threading.Lock()


# --------------------------------------------------------------------------- #
#  Cookie helpers
# --------------------------------------------------------------------------- #
def _host(url: str) -> str:
    return (urlparse(url).hostname or "").lower()


def _short(url: str) -> str:
    """'https://login.microsoftonline.com/x?y=1' -> 'login.microsoftonline.com'."""
    parsed = urlparse(url or "")
    return parsed.hostname or (url or "nothing")[:40] or "nothing"


def is_login_page(url: str) -> bool:
    """True when a URL is an identity provider's sign-in page, not the app."""
    u = (url or "").lower()
    return any(marker in u for marker in _LOGIN_HOSTS)


def _cookie_header(cookies: list[dict], host: str) -> str:
    """Join the cookies that belong to `host` into a 'a=1; b=2' header."""
    parts, seen = [], set()
    for c in cookies:
        name = c.get("name")
        if not name or name in seen:
            continue
        domain = (c.get("domain") or "").lstrip(".").lower()
        if domain and not (host == domain or host.endswith("." + domain)):
            continue
        seen.add(name)
        parts.append(f"{name}={c.get('value', '')}")
    return "; ".join(parts)


def _client_from_cookies(base_url: str, cookie_header: str, verify_ssl: bool):
    """Try the cookies against Jira. Returns (client, display_name) or (None, None)."""
    if "JSESSIONID" not in cookie_header:
        return None, None
    try:
        # A short timeout: this is a probe, and several of them run in a row.
        client = core.JiraClient(base_url, session_cookie=cookie_header,
                                 verify_ssl=verify_ssl, timeout=PROBE_TIMEOUT)
        who = client.verify_login()
        client.timeout = 30                  # back to normal for real work
        return client, who
    except Exception:  # noqa: BLE001 — an unusable cookie is not an error here
        return None, None


def _client_from_saved(base_url: str, verify_ssl: bool):
    """
    Sign in with the credentials we remembered.
    Returns (client, display_name, had_saved) — `had_saved` tells the caller
    whether there was anything to try, so a rejected password reads
    differently from never having saved one.
    """
    saved = jira_credentials.load(base_url)
    if not saved or not saved.get("password"):
        return None, None, False
    try:
        client = core.JiraClient(base_url, saved.get("username", ""),
                                 saved["password"], verify_ssl=verify_ssl,
                                 timeout=PROBE_TIMEOUT)
        who = client.verify_login()
        client.timeout = 30
        return client, who, True
    except Exception:  # noqa: BLE001 — stale password: fall through to the rest
        return None, None, True


def _existing_browser_sessions(base_url: str):
    """Yield (browser_name, cookie_header) for browsers signed in to this Jira."""
    try:
        import browser_cookie3  # noqa: WPS433
    except ImportError:
        return
    host = _host(base_url)
    loaders = ("edge", "chrome", "firefox", "brave", "chromium", "opera")
    for name in loaders:
        loader = getattr(browser_cookie3, name, None)
        if loader is None:
            continue
        try:
            jar = loader(domain_name=host)
        except Exception:  # noqa: BLE001 — locked/encrypted store, just move on
            continue
        header = _cookie_header(
            [{"name": c.name, "value": c.value, "domain": c.domain} for c in jar],
            host,
        )
        if header:
            yield name, header


# --------------------------------------------------------------------------- #
#  Browser automation
# --------------------------------------------------------------------------- #
def launch_browser(download_dir: Optional[str] = None, headless: bool = True,
                   page_load_strategy: str = "eager"):
    """
    Drive a browser on the shared profile — headless, so nothing pops up on
    screen. Because the profile persists, whatever signed in last time
    (Microsoft, Jira, the attendance portal) is still signed in, which is what
    makes running unattended possible. Raises RuntimeError with advice when no
    browser can be driven.
    """
    try:
        from selenium import webdriver  # noqa: WPS433
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.edge.options import Options as EdgeOptions
    except ImportError as exc:
        raise RuntimeError(
            "Automatic sign-in needs Selenium. Run:  pip install selenium"
        ) from exc

    os.makedirs(PROFILE_DIR, exist_ok=True)
    attempts = (
        ("Edge", EdgeOptions, webdriver.Edge),
        ("Chrome", ChromeOptions, webdriver.Chrome),
    )
    problems = []
    for label, options_cls, driver_cls in attempts:
        try:
            opts = options_cls()
            # A dedicated profile: keeps the Microsoft session between runs and
            # never collides with the browser the user already has open.
            opts.add_argument(f"--user-data-dir={os.path.join(PROFILE_DIR, label.lower())}")
            opts.add_argument("--no-first-run")
            opts.add_argument("--no-default-browser-check")
            opts.add_experimental_option("excludeSwitches", ["enable-automation"])
            opts.add_experimental_option("useAutomationExtension", False)
            # 'eager' returns as soon as the DOM is ready; 'none' returns at
            # once, which suits a single-page app we are going to poll anyway.
            opts.page_load_strategy = page_load_strategy
            if headless:
                # No window at any point; the size still matters because the
                # page is measured for visible elements.
                opts.add_argument("--headless=new")
                opts.add_argument("--disable-gpu")
                opts.add_argument("--window-size=1440,1000")
            if download_dir:
                opts.add_experimental_option("prefs", {
                    "download.default_directory": download_dir,
                    "download.prompt_for_download": False,
                    "download.directory_upgrade": True,
                    "safebrowsing.enabled": True,
                })
            return driver_cls(options=opts)
        except Exception as exc:  # noqa: BLE001
            problems.append(f"{label}: {exc}")
    raise RuntimeError("Couldn't start a browser. " + " | ".join(problems))


# --------------------------------------------------------------------------- #
#  The login job
# --------------------------------------------------------------------------- #
class SsoLogin:
    """Runs one sign-in attempt in the background and reports its progress."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None
        self._cancel = threading.Event()
        self._launched = False
        self._state = {"state": "idle", "message": "", "user": None,
                       "can_open_window": False}

    # ---- state ---------------------------------------------------------- #
    def _set(self, state: str, message: str, user: Optional[str] = None,
             can_open_window: bool = False) -> None:
        with self._lock:
            self._state = {"state": state, "message": message, "user": user,
                           "can_open_window": can_open_window}

    def status(self) -> dict:
        with self._lock:
            state = dict(self._state)
        # Same guarantee as the attendance fetch: never spin on a dead worker.
        if state["state"] == "working" and self._launched and not self.busy():
            state.update(state="error", can_open_window=True,
                         message="The sign-in stopped unexpectedly. Try again.")
            with self._lock:
                self._state = dict(state)
        return state

    def busy(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    def cancel(self) -> None:
        self._cancel.set()
        self._set("idle", "")

    # ---- run ------------------------------------------------------------- #
    def start(self, base_url: str, verify_ssl: bool,
              on_success: Callable[[object, str], None],
              interactive: bool = False) -> tuple:
        """
        Kick off a sign-in. Silent (no window) by default; `interactive` opens
        a real window, which is the only way to get past a Microsoft password
        prompt. Returns (ok, error_message).
        """
        if self.busy():
            if not self._cancel.is_set():
                return True, ""      # already running; the UI just keeps polling
            self._thread.join(timeout=5)   # let a cancelled run close its browser
            if self.busy():
                return False, "Still closing the previous sign-in — try again."
        self._cancel.clear()
        self._launched = False
        self._set("working", "Starting…")
        self._thread = threading.Thread(
            target=self._run, args=(base_url, verify_ssl, on_success, interactive),
            daemon=True,
        )
        self._thread.start()
        self._launched = True
        return True, ""

    def _run(self, base_url: str, verify_ssl: bool, on_success,
             interactive: bool = False) -> None:
        base_url = base_url.rstrip("/")
        host = _host(base_url)
        had_saved = False
        try:
            # --- 1. the sign-in we already remember (instant, one request) -- #
            if not interactive:
                self._set("working", "Signing you in…")
                client, who, had_saved = _client_from_saved(base_url, verify_ssl)
                if client:
                    on_success(client, who)
                    return self._set("done", f"Signed in as {who}", who)

            # --- 2. a session token from a browser already signed in -------- #
            if not interactive:
                for name, header in _existing_browser_sessions(base_url):
                    if self._cancel.is_set():
                        return self._set("idle", "")
                    self._set("working", f"Trying the session in {name.title()}…")
                    client, who = _client_from_cookies(base_url, header, verify_ssl)
                    if client:
                        on_success(client, who)
                        return self._set("done", f"Signed in as {who}", who)

            # --- 3. nothing saved, and this Jira has no SSO to walk --------- #
            # Walking a password-only Jira headlessly can only land on its login
            # form, so don't spend 45 seconds proving it.
            if not interactive and any(h in host for h in PASSWORD_ONLY_HOSTS):
                return self._set(
                    "error",
                    (_SAVED_REJECTED if had_saved else _NEED_CREDENTIALS).format(where=host),
                    None, False)

            # --- otherwise: walk the SSO redirect ourselves ----------------- #
            self._set("working", "Opening a sign-in window…" if interactive
                      else "Signing in with Microsoft…")
            if not BROWSER_LOCK.acquire(timeout=180):
                return self._set("error", "The browser is busy with another step.")
            driver = None
            try:
                driver = launch_browser(headless=not interactive)
                try:
                    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
                except Exception:  # noqa: BLE001
                    pass
                try:
                    driver.get(f"{base_url}/secure/Dashboard.jspa")
                except Exception:  # noqa: BLE001
                    pass          # slow page: keep going, it loads underneath us
                url = ""
                started = time.time()
                deadline = started + (WINDOW_TIMEOUT if interactive
                                      else SILENT_TIMEOUT)
                if interactive:
                    self._set("working", "Sign in with Microsoft in the window "
                                         "that just opened — this page continues "
                                         "by itself.")
                while time.time() < deadline:
                    if self._cancel.is_set():
                        return self._set("idle", "")
                    try:
                        cookies = driver.get_cookies()
                        url = driver.current_url or ""
                    except Exception:  # noqa: BLE001 — the browser went away
                        return self._set("error", "The browser stopped before "
                                                  "Jira signed you in.", None, True)
                    client, who = _client_from_cookies(
                        base_url, _cookie_header(cookies, host), verify_ssl)
                    if client:
                        on_success(client, who)
                        return self._set("done", f"Signed in as {who}", who)
                    if not interactive:
                        # Nothing can be clicked in the background, so stop as
                        # soon as we know a human is needed — but give any SSO
                        # redirect a moment to happen first.
                        settled = time.time() - started > _SSO_GRACE
                        try:
                            password_form = settled and driver.execute_script(
                                _JS_JIRA_LOGIN_FORM)
                        except Exception:  # noqa: BLE001
                            password_form = False
                        if password_form:
                            # A password form is the end of the road for a
                            # background sign-in: ask for credentials to save.
                            return self._set(
                                "error", _NEED_CREDENTIALS.format(where=_short(url)),
                                None, False)
                        if is_login_page(url):
                            return self._set(
                                "error",
                                _NOT_SIGNED_IN.format(url=base_url, where=_short(url)),
                                None, True)
                        self._set("working", f"Signing in with Microsoft… ({_short(url)})")
                    time.sleep(POLL_SECONDS)
                self._set("error",
                          _NOT_SIGNED_IN.format(url=base_url, where=_short(url)),
                          None, True)
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:  # noqa: BLE001
                        pass
                BROWSER_LOCK.release()
        except Exception as exc:  # noqa: BLE001
            self._set("error", str(exc))


MANAGER = SsoLogin()
