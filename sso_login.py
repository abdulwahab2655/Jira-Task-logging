#!/usr/bin/env python3
"""
sso_login.py — one-click sign-in for the Jira timesheet web UI.

The page shows one button. Clicking it signs you in the way the attendance
portal does: if a live session can be found you are simply in, and if not, a
sign-in window opens, you sign in there as you normally would, and the window
closes by itself the moment Jira hands out a session.

In order:

  1. A username + password (or PAT) saved from an earlier run — one REST call,
     effectively instant, nothing on screen.
  2. A browser on this machine that already holds a live Jira session; we
     borrow those cookies.
  3. Our own browser profile, headless. A sign-in window leaves its cookies
     there, so every run after the first one is a single click and no window.
  4. The sign-in window itself. Whatever the login page is — the stock
     Atlassian form, Microsoft, Okta, a second factor — it is a real browser,
     so a person can get through it. We only watch for the session that comes
     out the other end.

The browser we drive keeps its own profile folder next to this file, which is
why step 3 works at all: the session established in the window last time is
still there.

Optional dependencies (both degrade gracefully):
    pip install selenium        # drives the sign-in window  (the main path)
    pip install browser_cookie3 # reuse an already-signed-in browser (shortcut)
"""

from __future__ import annotations

import os
import threading
import time
from contextlib import contextmanager
from typing import Callable, Optional
from urllib.parse import urlparse

import jira_credentials
import jira_logging_utility as core

HERE = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.join(HERE, ".sso-browser-profile")

WINDOW_POLL = 0.6              # how often the open window is checked, so it
                               # closes promptly once the session appears
PROFILE_TIMEOUT = 14           # looking for a session we already have
WINDOW_TIMEOUT = 300           # a human signing in, in the window
PAGE_LOAD_TIMEOUT = 25         # never block forever inside driver.get()
PROBE_TIMEOUT = 8              # per cookie-probe request

# Jira Server/DC's own login page, which is also what redirects to whatever
# identity provider the instance uses. os_destination lands you on the
# dashboard once you are through, which is when the cookies we want exist.
LOGIN_PATH = "/login.jsp?os_destination=%2Fsecure%2FDashboard.jspa"

# Roughly the shape of an OAuth popup, so it reads as a sign-in window rather
# than a browser someone left open.
POPUP_SIZE = (560, 780)

_WINDOW_TIMED_OUT = (
    "The sign-in window was open for {mins} minutes without Jira handing out a "
    "session. Click Sign in with Jira to try again, or use a password below."
)

_WINDOW_CLOSED = (
    "That window closed before Jira signed you in. Click Sign in with Jira to "
    "open it again — or use a password below."
)

# The stock Atlassian login form — or, on the anonymous dashboard, the plain
# "Log in" link that stands in for it. Either means: this profile is signed
# out, so there is nothing for a background check to find.
#
# This Jira spells the fields username-field / password-field; older ones use
# login-form-username. Both are here, and so is the password box itself, which
# is the one thing every login page has and no signed-in page does.
_JS_JIRA_LOGIN_FORM = """
return !!(document.getElementById('username-field') ||
          document.getElementById('password-field') ||
          document.getElementById('login-form-username') ||
          document.querySelector('form#login-form input[type=password]') ||
          document.querySelector("a[href*='login.jsp']"));
"""

# A small courtesy in the window: put the username we already know in the box
# and leave the cursor in the password field. The password is never filled —
# typing it is the whole reason the window is open.
_JS_PREFILL_USERNAME = """
var name = arguments[0];
var box = document.getElementById('username-field') ||
          document.getElementById('login-form-username') ||
          document.querySelector("form#login-form input[name='username']") ||
          document.querySelector("form#login-form input[name='os_username']");
if (!box || box.value) return false;
box.value = name;
box.dispatchEvent(new Event('input', {bubbles: true}));
var pw = document.getElementById('password-field') ||
         document.getElementById('login-form-password') ||
         document.querySelector("form#login-form input[type=password]");
if (pw) { pw.focus(); }
return true;
"""

# One profile means one browser at a time: the sign-in and the attendance fetch
# take turns rather than colliding over a locked user-data-dir.
BROWSER_LOCK = threading.Lock()


class BrowserUnavailable(RuntimeError):
    """
    No browser can be driven on this machine.

    Its own class because it is the one failure a sign-in window cannot fix:
    the window *is* the browser. The UI offers the password form instead.
    """


@contextmanager
def browser_lock(timeout: int = 180):
    """Hold the single-browser lock, or say what is holding it up."""
    if not BROWSER_LOCK.acquire(timeout=timeout):
        raise RuntimeError("The browser is busy with another step.")
    try:
        yield
    finally:
        BROWSER_LOCK.release()


# --------------------------------------------------------------------------- #
#  Cookie helpers
# --------------------------------------------------------------------------- #
def _host(url: str) -> str:
    return (urlparse(url).hostname or "").lower()


def _short(url: str) -> str:
    """'https://login.microsoftonline.com/x?y=1' -> 'login.microsoftonline.com'."""
    parsed = urlparse(url or "")
    return parsed.hostname or (url or "nothing")[:40] or "nothing"


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
        parts.append("{}={}".format(name, c.get("value", "")))
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
    Sign in with the credentials we remembered, if there are any.
    Returns (client, display_name, had_saved) — had_saved tells the caller
    whether there was anything to try at all.
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


def _profile_exists() -> bool:
    """
    Whether we have a browser profile worth searching.

    On a first run there is nothing in it, and starting a browser to prove that
    costs the eight seconds standing between the click and the sign-in window.
    """
    return any(os.path.isdir(os.path.join(PROFILE_DIR, name))
               for name in ("edge", "chrome"))


def _saved_username(base_url: str) -> str:
    """The username we already know, for the window to start filled in."""
    try:
        return (jira_credentials.load(base_url) or {}).get("username") or ""
    except Exception:  # noqa: BLE001 — a courtesy, never a reason to stop
        return ""


# --------------------------------------------------------------------------- #
#  Browser automation
# --------------------------------------------------------------------------- #
def launch_browser(download_dir: Optional[str] = None, headless: bool = True,
                   page_load_strategy: str = "eager",
                   popup_url: Optional[str] = None,
                   window_size: Optional[tuple] = None):
    """
    Drive a browser on the shared profile. Headless by default, so nothing
    appears on screen; pass headless=False with a popup_url for the sign-in
    window, which opens chromeless at that address.

    Because the profile persists, whatever signed in last time (Jira,
    Microsoft, the attendance portal) is still signed in — which is what makes
    running unattended possible. Raises RuntimeError with advice when no
    browser can be driven.
    """
    try:
        from selenium import webdriver  # noqa: WPS433
        from selenium.webdriver.chrome.options import Options as ChromeOptions
        from selenium.webdriver.edge.options import Options as EdgeOptions
    except ImportError as exc:
        raise BrowserUnavailable(
            "The sign-in window needs Selenium. Run:  pip install selenium — "
            "or use the password form under the button, which does not."
        ) from exc

    os.makedirs(PROFILE_DIR, exist_ok=True)
    attempts = (
        ("Edge", EdgeOptions, webdriver.Edge),
        ("Chrome", ChromeOptions, webdriver.Chrome),
    )
    # App mode is what makes the sign-in window look like a sign-in window
    # rather than a browser. Not every driver build tolerates it, so a plain
    # window is the fallback.
    app_modes = (True, False) if (popup_url and not headless) else (False,)
    problems = []
    for label, options_cls, driver_cls in attempts:
        for app_mode in app_modes:
            try:
                opts = options_cls()
                # A dedicated profile: keeps the sessions between runs and
                # never collides with the browser the user already has open.
                opts.add_argument(
                    "--user-data-dir=" + os.path.join(PROFILE_DIR, label.lower()))
                opts.add_argument("--no-first-run")
                opts.add_argument("--no-default-browser-check")
                opts.add_experimental_option("excludeSwitches", ["enable-automation"])
                opts.add_experimental_option("useAutomationExtension", False)
                # 'eager' returns as soon as the DOM is ready; 'none' returns at
                # once, which suits a page we are going to poll anyway.
                opts.page_load_strategy = page_load_strategy
                width, height = window_size or (
                    POPUP_SIZE if popup_url and not headless else (1440, 1000))
                opts.add_argument("--window-size={},{}".format(width, height))
                if app_mode:
                    opts.add_argument("--app=" + popup_url)
                if headless:
                    # No window at any point; the size still matters because
                    # the page is measured for visible elements.
                    opts.add_argument("--headless=new")
                    opts.add_argument("--disable-gpu")
                if download_dir:
                    opts.add_experimental_option("prefs", {
                        "download.default_directory": download_dir,
                        "download.prompt_for_download": False,
                        "download.directory_upgrade": True,
                        "safebrowsing.enabled": True,
                    })
                driver = driver_cls(options=opts)
                # App mode navigates as the window opens. Saying so here stops
                # the caller loading the same page a second time: while that
                # reload is in flight the document is empty, and an empty
                # document is indistinguishable from a page with no login form.
                driver.opened_at_url = popup_url if app_mode else None
                return driver
            except Exception as exc:  # noqa: BLE001
                problems.append("{}{}: {}".format(
                    label, " (app window)" if app_mode else "", exc))
    raise BrowserUnavailable(
        "Couldn't start a browser, so there is nowhere to sign in. Check Edge "
        "or Chrome is installed, or use the password form under the button. "
        + " | ".join(problems))


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
                       "stage": "", "can_open_window": False}

    # ---- state ---------------------------------------------------------- #
    def _set(self, state: str, message: str, user: Optional[str] = None,
             can_open_window: bool = False, stage: str = "") -> None:
        with self._lock:
            self._state = {"state": state, "message": message, "user": user,
                           "stage": stage, "can_open_window": can_open_window}

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
        Kick off a sign-in. By default it looks for a session you already have
        and only opens the window when it has to; interactive skips the looking
        and opens the window straight away. Returns (ok, error_message).
        """
        if self.busy():
            if not self._cancel.is_set():
                return True, ""      # already running; the UI just keeps polling
            self._thread.join(timeout=5)   # let a cancelled run close its browser
            if self.busy():
                return False, "Still closing the previous sign-in — try again."
        self._cancel.clear()
        self._launched = False
        self._set("working", "Starting…", stage="checking")
        self._thread = threading.Thread(
            target=self._run, args=(base_url, verify_ssl, on_success, interactive),
            daemon=True,
        )
        self._thread.start()
        self._launched = True
        return True, ""

    # ---- the pieces of a sign-in ----------------------------------------- #
    def _adopt(self, client, who: str, on_success) -> None:
        """Take a working session as ours and report the sign-in done."""
        on_success(client, who)
        self._set("done", "Signed in as {}".format(who), who, stage="done")

    def _cancelled(self) -> bool:
        if self._cancel.is_set():
            self._set("idle", "")
            return True
        return False

    def _profile_session(self, base_url: str, verify_ssl: bool):
        """
        The Jira session our own browser profile holds, with nothing on screen.
        Returns (client, display_name) or (None, None).

        This is the step the window pays for: a sign-in done there leaves its
        cookies in this profile, so the next run is one click and silence.
        """
        host = _host(base_url)
        with browser_lock():
            driver = None
            try:
                driver = launch_browser(headless=True, page_load_strategy="none")
                try:
                    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
                    driver.get(base_url + "/secure/Dashboard.jspa")
                except Exception:  # noqa: BLE001 — it loads underneath us
                    pass
                deadline = time.time() + PROFILE_TIMEOUT
                while time.time() < deadline:
                    if self._cancel.is_set():
                        return None, None
                    try:
                        cookies = driver.get_cookies()
                    except Exception:  # noqa: BLE001 — the browser went away
                        return None, None
                    client, who = _client_from_cookies(
                        base_url, _cookie_header(cookies, host), verify_ssl)
                    if client:
                        return client, who
                    try:
                        if driver.execute_script(_JS_JIRA_LOGIN_FORM):
                            return None, None   # signed out: a person is needed
                    except Exception:  # noqa: BLE001
                        pass
                    time.sleep(0.5)
                return None, None
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:  # noqa: BLE001
                        pass

    def _window_session(self, base_url: str, verify_ssl: bool, on_success) -> None:
        """
        Open the sign-in window and wait for a session to come out of it.

        We do not drive that page: it may be Jira's own form, Microsoft, or a
        second factor, and all of those are a person's job. All we do is watch
        the cookies — and the moment they work, the window closes and the app
        is signed in.
        """
        host = _host(base_url)
        username = _saved_username(base_url)
        url = base_url + LOGIN_PATH
        with browser_lock():
            driver = None
            try:
                self._set("working", "Opening the sign-in window…", stage="window")
                driver = launch_browser(headless=False, page_load_strategy="none",
                                        popup_url=url)
                try:
                    driver.set_page_load_timeout(PAGE_LOAD_TIMEOUT)
                except Exception:  # noqa: BLE001
                    pass
                if getattr(driver, "opened_at_url", None) != url:
                    try:
                        driver.get(url)
                    except Exception:  # noqa: BLE001 — it loads underneath us
                        pass
                self._set("working",
                          "Sign in in the window that just opened. It closes by "
                          "itself as soon as Jira lets you in.", stage="window")
                deadline = time.time() + WINDOW_TIMEOUT
                prefilled = False
                while time.time() < deadline:
                    if self._cancel.is_set():
                        return self._set("idle", "")
                    try:
                        cookies = driver.get_cookies()
                    except Exception:  # noqa: BLE001
                        break            # the window is gone — see why below
                    client, who = _client_from_cookies(
                        base_url, _cookie_header(cookies, host), verify_ssl)
                    if client:
                        return self._adopt(client, who, on_success)
                    if username and not prefilled:
                        try:
                            prefilled = bool(driver.execute_script(
                                _JS_PREFILL_USERNAME, username))
                        except Exception:  # noqa: BLE001
                            pass
                    time.sleep(WINDOW_POLL)
                else:
                    return self._set("error", _WINDOW_TIMED_OUT.format(
                        mins=WINDOW_TIMEOUT // 60), None, True)
            finally:
                if driver is not None:
                    try:
                        driver.quit()
                    except Exception:  # noqa: BLE001
                        pass

        # The window was closed by hand. If the sign-in went through just
        # before that, its cookies are in the profile — so closing the window
        # yourself works just as well as letting it close itself.
        if self._cancelled():
            return None
        self._set("working", "Checking whether that signed you in…",
                  stage="checking")
        client, who = self._profile_session(base_url, verify_ssl)
        if client:
            return self._adopt(client, who, on_success)
        return self._set("error", _WINDOW_CLOSED, None, True)

    def _run(self, base_url: str, verify_ssl: bool, on_success,
             interactive: bool = False) -> None:
        base_url = base_url.rstrip("/")
        try:
            if not interactive:
                # --- 1. the sign-in we already remember (one request) ------ #
                self._set("working", "Signing you in…", stage="checking")
                client, who, _had_saved = _client_from_saved(base_url, verify_ssl)
                if client:
                    return self._adopt(client, who, on_success)
                if self._cancelled():
                    return None

                # --- 2. a session from a browser already signed in --------- #
                for name, header in _existing_browser_sessions(base_url):
                    if self._cancelled():
                        return None
                    self._set("working",
                              "Trying the session in {}…".format(name.title()),
                              stage="checking")
                    client, who = _client_from_cookies(base_url, header, verify_ssl)
                    if client:
                        return self._adopt(client, who, on_success)

                # --- 3. the session the window left here last time --------- #
                if _profile_exists():
                    self._set("working", "Looking for a session you already have…",
                              stage="checking")
                    client, who = self._profile_session(base_url, verify_ssl)
                    if client:
                        return self._adopt(client, who, on_success)
                    if self._cancelled():
                        return None

            # --- 4. ask the person: open the sign-in window ---------------- #
            return self._window_session(base_url, verify_ssl, on_success)
        except BrowserUnavailable as exc:
            # Nothing to reopen: the window itself is what is missing.
            self._set("error", str(exc), None, False)
        except Exception as exc:  # noqa: BLE001
            self._set("error", str(exc), None, True)
        return None


MANAGER = SsoLogin()
