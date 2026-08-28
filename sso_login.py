#!/usr/bin/env python3
"""
sso_login.py — one-click sign-in for the Jira timesheet web UI.

The page shows one button. Clicking it signs you in the way the attendance
portal does: if a live session can be found you are simply in, and if not, a
sign-in window opens, you sign in there as you normally would, and the window
closes by itself the moment Jira hands out a session.

The window comes first, because opening a browser is the slow part of any
sign-in that needs one — everything else runs beside it:

  * A saved username + password (or PAT) gets a short head start: one REST
     call, and if it lands nothing opens at all.
  * The window opens on Jira's login page, with the saved username and
     password already in the boxes, so there is one button left to press.
     Whatever that page is — the stock Atlassian form, Microsoft, Okta, a
     second factor — it is a real browser, so a person can get through it.
  * While it is open we keep looking for a session that would make it
     unnecessary: the browsers on this machine, and the saved password if it
     was still answering. The first one to work closes the window.
  * A password typed into Jira's own form is saved when it works, so the next
     run takes the first path and opens nothing.

The browser we drive keeps its own profile folder next to this file. A session
established in the window is still there next time, so the window opens
already signed in and closes itself in about a second.

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

WINDOW_POLL = 0.35             # how often the open window is checked, so it
                               # closes promptly once the session appears
QUICK_WAIT = 1.6               # head start for the saved password, before a
                               # window opens that it may make unnecessary
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

# The window opens with whatever we already know in it: the username, and the
# password too when one is saved. That leaves one button to press — so the
# cursor goes to it, or to the password box when there is nothing to put there.
_JS_PREFILL_LOGIN = """
var user = arguments[0], pass = arguments[1];
var form = document.getElementById('login-form');
var box = document.getElementById('username-field') ||
          document.getElementById('login-form-username') ||
          (form && form.querySelector("input[name='username'],input[name='os_username']"));
if (!box) return false;
var pw = document.getElementById('password-field') ||
         document.getElementById('login-form-password') ||
         (form && form.querySelector('input[type=password]'));
var did = false;
if (user && !box.value) {
  box.value = user;
  box.dispatchEvent(new Event('input', {bubbles: true}));
  did = true;
}
if (pw && pass && !pw.value) {
  pw.value = pass;
  pw.dispatchEvent(new Event('input', {bubbles: true}));
  did = true;
}
var go = form && form.querySelector("input[type=submit],button[type=submit],#login");
if (pw && !pw.value) { pw.focus(); }
else if (go && go.focus) { go.focus(); }
return did;
"""

# What is in Jira's own login form right now. Read only while the window is on
# the Jira host itself, and used for one thing: remembering a password that
# worked, so the next run signs in without opening anything. This is the same
# password the old sign-in form on the page used to collect and save.
_JS_READ_LOGIN = """
var form = document.getElementById('login-form');
if (!form) return null;
var box = document.getElementById('username-field') ||
          document.getElementById('login-form-username') ||
          form.querySelector("input[name='username'],input[name='os_username']");
var pw = document.getElementById('password-field') ||
         document.getElementById('login-form-password') ||
         form.querySelector('input[type=password]');
if (!box || !pw || !box.value || !pw.value) return null;
return {u: box.value, p: pw.value};
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


class _SessionWatch:
    """
    Watch a driven browser for a Jira session without hammering Jira.

    The obvious loop — "every tick, try the cookies" — costs a real request
    every tick, because Jira hands out an anonymous JSESSIONID the moment its
    login page loads. "Has a session cookie" is therefore true from the start
    and cannot be the trigger. What actually changes when someone signs in is
    the cookie itself (Jira rotates the session id) and the page they are on.

    So: a request when either of those changes, and otherwise one every few
    seconds in case some Jira signs you in without rotating anything. Typing a
    password now costs nothing, and the sign-in is noticed sooner than before.
    """

    HEARTBEAT = 3.0

    def __init__(self, base_url: str, host: str, verify_ssl: bool) -> None:
        self.base_url = base_url
        self.host = host
        self.verify_ssl = verify_ssl
        self._seen = None          # the cookie header we last asked about
        self._url = None
        self._asked = 0.0
        self.requests = 0          # how many probes we actually spent

    def look(self, cookies: list, url: str = "") -> tuple:
        """(client, display_name), or (None, None) — including "not worth asking"."""
        header = _cookie_header(cookies, self.host)
        now = time.time()
        if (header == self._seen and url == self._url
                and now - self._asked < self.HEARTBEAT):
            return None, None
        self._seen, self._url, self._asked = header, url, now
        self.requests += 1
        return _client_from_cookies(self.base_url, header, self.verify_ssl)


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


def _saved_login(base_url: str) -> tuple:
    """(username, password) we already know, for the window to start filled in."""
    try:
        saved = jira_credentials.load(base_url) or {}
        return saved.get("username") or "", saved.get("password") or ""
    except Exception:  # noqa: BLE001 — a courtesy, never a reason to stop
        return "", ""


def _remember_login(base_url: str, username: str, password: str) -> str:
    """Save a sign-in that worked. Returns where it went, or ''."""
    try:
        return jira_credentials.save(base_url, username, password) or ""
    except Exception:  # noqa: BLE001 — signed in is signed in; saving is extra
        return ""


def _remember_verified_login(base_url: str, verify_ssl: bool,
                             username: str, password: str) -> str:
    """
    Save a sign-in typed into Jira's form — but only once it is proved.

    That form is read by polling, so what we are holding may be a password
    caught between keystrokes. Saving a truncated one would make the next run
    sign in with a wrong password, and a few of those is how Jira starts asking
    for a CAPTCHA. One request settles it, and it is the same request the next
    run will make anyway.
    """
    if not (username and password):
        return ""
    try:
        client = core.JiraClient(base_url, username, password,
                                 verify_ssl=verify_ssl, timeout=PROBE_TIMEOUT)
        client.verify_login()
    except Exception:  # noqa: BLE001 — half a password is not worth keeping
        return ""
    return _remember_login(base_url, username, password)


class _Probe:
    """
    The sign-ins that need no window, running while one opens.

    Two of them, in the order they answer: the password we saved (one request)
    and then the browsers on this machine (several cookie stores, and a locked
    one can take seconds — which is exactly why it does not hold the window up
    any more). `saved_done` is what the caller waits on for its head start;
    `result` is filled the moment anything works, and the window loop watches
    for it.
    """

    def __init__(self, base_url: str, verify_ssl: bool) -> None:
        self.base_url = base_url
        self.verify_ssl = verify_ssl
        self.result = None                  # (client, display_name)
        self.saved_done = threading.Event()
        self.stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def start(self):
        self._thread.start()
        return self

    def _run(self) -> None:
        try:
            client, who, _had = _client_from_saved(self.base_url, self.verify_ssl)
            if client:
                self.result = (client, who)
        except Exception:  # noqa: BLE001
            pass
        finally:
            self.saved_done.set()
        if self.result or self.stop.is_set():
            return
        try:
            for _name, header in _existing_browser_sessions(self.base_url):
                if self.stop.is_set():
                    return
                client, who = _client_from_cookies(self.base_url, header,
                                                   self.verify_ssl)
                if client:
                    self.result = (client, who)
                    return
        except Exception:  # noqa: BLE001
            pass


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
                       "stage": "", "saved": "", "can_open_window": False}

    # ---- state ---------------------------------------------------------- #
    def _set(self, state: str, message: str, user: Optional[str] = None,
             can_open_window: bool = False, stage: str = "",
             saved: str = "") -> None:
        with self._lock:
            self._state = {"state": state, "message": message, "user": user,
                           "stage": stage, "saved": saved,
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
    def _adopt(self, client, who: str, on_success, saved: str = "") -> None:
        """Take a working session as ours and report the sign-in done."""
        on_success(client, who)
        self._set("done", "Signed in as {}".format(who), who, stage="done",
                  saved=saved)

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
        watch = _SessionWatch(base_url, host, verify_ssl)
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
                    client, who = watch.look(cookies)
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

    def _window_session(self, base_url: str, verify_ssl: bool, on_success,
                        probe=None) -> None:
        """
        Open the sign-in window and wait for a session to come out of it.

        We do not drive that page: it may be Jira's own form, Microsoft, or a
        second factor, and all of those are a person's job. What we do is fill
        in what we already know, watch the cookies, and keep an eye on the
        `probe` running beside us — whichever produces a session first closes
        the window.

        A password typed into Jira's own form is remembered when it works, so
        the next run is a single request and nothing opens.
        """
        host = _host(base_url)
        username, password = _saved_login(base_url)
        url = base_url + LOGIN_PATH
        typed = None            # the last complete sign-in seen in the form
        watch = _SessionWatch(base_url, host, verify_ssl)
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
                    # A session found without the window ends it too: the
                    # saved password may simply have been slower than the
                    # browser was to open.
                    if probe is not None and probe.result:
                        client, who = probe.result
                        return self._adopt(client, who, on_success)
                    # Only the cookies say whether the window is still there.
                    # current_url throws while a page is navigating, and
                    # treating that as "gone" used to close the window in the
                    # middle of a sign-in and blame the user for it.
                    try:
                        cookies = driver.get_cookies()
                    except Exception:  # noqa: BLE001
                        break            # the window is gone — see why below
                    try:
                        here = _host(driver.current_url or "")
                    except Exception:  # noqa: BLE001 — mid-navigation
                        here = ""
                    client, who = watch.look(cookies, here)
                    if client:
                        saved_to = ""
                        if typed:
                            saved_to = _remember_verified_login(
                                base_url, verify_ssl, *typed)
                        return self._adopt(client, who, on_success, saved_to)
                    # Only on Jira's own host, and only its own login form: an
                    # identity provider's password is not ours to keep.
                    if here == host:
                        if not prefilled and (username or password):
                            try:
                                prefilled = bool(driver.execute_script(
                                    _JS_PREFILL_LOGIN, username, password))
                            except Exception:  # noqa: BLE001
                                pass
                        try:
                            seen = driver.execute_script(_JS_READ_LOGIN)
                        except Exception:  # noqa: BLE001
                            seen = None
                        if seen and seen.get("u") and seen.get("p"):
                            typed = (seen["u"], seen["p"])
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
        if probe is not None and probe.result:
            client, who = probe.result
            return self._adopt(client, who, on_success)
        self._set("working", "Checking whether that signed you in…",
                  stage="checking")
        client, who = self._profile_session(base_url, verify_ssl)
        if client:
            return self._adopt(client, who, on_success,
                               _remember_verified_login(base_url, verify_ssl, *typed)
                               if typed else "")
        return self._set("error", _WINDOW_CLOSED, None, True)

    def _run(self, base_url: str, verify_ssl: bool, on_success,
             interactive: bool = False) -> None:
        base_url = base_url.rstrip("/")
        probe = None
        try:
            if not interactive:
                # The saved password gets a head start measured in one request:
                # if it answers, nothing opens. If it is slow, it keeps going
                # in the background and the window closes when it lands.
                self._set("working", "Signing you in…", stage="checking")
                probe = _Probe(base_url, verify_ssl).start()
                probe.saved_done.wait(QUICK_WAIT)
                if probe.result:
                    client, who = probe.result
                    return self._adopt(client, who, on_success)
                if self._cancelled():
                    return None

            # Straight to the window. It is the slow part, so it is not queued
            # behind anything, and the probe keeps looking while it is open.
            return self._window_session(base_url, verify_ssl, on_success, probe)
        except BrowserUnavailable as exc:
            # Nothing to reopen: the window itself is what is missing.
            self._set("error", str(exc), None, False)
        except Exception as exc:  # noqa: BLE001
            self._set("error", str(exc), None, True)
        finally:
            if probe is not None:
                probe.stop.set()
        return None


MANAGER = SsoLogin()
