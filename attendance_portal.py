#!/usr/bin/env python3
"""
attendance_portal.py — pull the attendance sheet straight from the portal.

The portal is a single-page app talking to a REST API, so we talk to that API
rather than scraping the screen. A headless browser opens the portal purely to
mint a session: if it lands on the portal's sign-in page it presses the
'Sign in with Microsoft' control (a custom <i2c-button>, so it is found by its
text) and the Microsoft session already in our shared profile completes the
round trip silently. We then lift the bearer token out of localStorage, and the
rest is two plain requests:

    GET /api/v1/admin/releases/…/withExceptionCutOff   -> release -> date range
    GET /api/v1/employee/attendances/getAttendanceAndSummary?startDate&endDate

The second returns one record per day with `totalHours` as a number and
explicit leave/holiday/day-off flags, which beats reading '9h 26m' off a grid.

Nothing appears on screen. Nothing in the background can type a password, so if
Microsoft actually stops for one we give up quickly and say so, rather than
sitting on a spinner.
"""

from __future__ import annotations

import base64
import json
import os
import threading
import time
from typing import Callable, Optional

import requests

import jira_logging_utility as core
import sso_login

PORTAL_URL = "https://attendance.i2cinc.com/employee/attendance"
API_BASE = "https://attendance-server-pilot.i2cinc.com/api/v1"
RELEASES_URL = f"{API_BASE}/admin/releases/engineering-releases-current-and-past/withExceptionCutOff"
ATTENDANCE_URL = f"{API_BASE}/employee/attendances/getAttendanceAndSummary"

HERE = os.path.dirname(os.path.abspath(__file__))

TOKEN_TIMEOUT = 60         # a silent SSO round trip, not a human typing
WINDOW_TIMEOUT = 300       # a human signing in, when they ask for a window
API_TIMEOUT = 60
POLL = 0.35                # the sign-in lands in a second or two, so look often
LOGIN_GRACE = 20           # how long the sign-in button gets to render
CLICK_COOLDOWN = 2.5       # don't hammer the button while the SPA re-renders

_NOT_SIGNED_IN = (
    "The attendance portal isn't signed in here yet. Open "
    "https://attendance.i2cinc.com/employee/attendance and sign in with "
    "Microsoft, then click Fetch attendance again — or open a sign-in window "
    "from here. After that it stays signed in and runs in the background."
)


# --------------------------------------------------------------------------- #
#  The portal's API
# --------------------------------------------------------------------------- #
def _headers(token: str, session_id: str) -> dict:
    return {"Accept": "application/json, text/plain, */*",
            "Authorization": f"Bearer {token}",
            "SessionId": session_id or ""}


def _get_json(url: str, headers: dict, params: Optional[dict] = None) -> dict:
    r = requests.get(url, headers=headers, params=params, timeout=API_TIMEOUT)
    if r.status_code == 401:
        raise RuntimeError("The portal rejected the session (401). Try again.")
    if r.status_code != 200:
        raise RuntimeError(f"{url.rsplit('/', 1)[-1]} failed (HTTP {r.status_code}).")
    return r.json()


def release_range(headers: dict, release: str) -> dict:
    """
    Map a release number to its dates: {'name', 'start', 'end', 'release_end'}.

    The API gives three ends for a release, e.g. for 26.08:
        codeFreezeDate 2026-08-18   the release end the portal displays
        rdDate         2026-08-23
        endDate        2026-08-23   stretched so exceptions can still be filed
    We *fetch* through the widest (`endDate`) so no worked day is missing, and
    report `codeFreezeDate` — the range the portal itself shows, 22 Jul → 18 Aug.
    """
    wanted = core.parse_release_number(release)
    if not wanted:
        raise RuntimeError(f"Couldn't read a release number from '{release}'.")
    data = _get_json(RELEASES_URL, headers).get("data") or []
    for rel in data:
        if core.parse_release_number(rel.get("number") or rel.get("name")) == wanted:
            return {"name": rel.get("number") or rel.get("name"),
                    "start": rel.get("startDate"),
                    "end": rel.get("endDate") or rel.get("codeFreezeDate"),
                    "release_end": rel.get("codeFreezeDate") or rel.get("endDate")}
    known = ", ".join(str(r.get("number")) for r in data[:8])
    raise RuntimeError(f"The portal has no release {release}. It lists: {known}.")


def _leave_label(rec: dict) -> str:
    """
    The leave *type* lives in `remarks` ('Casual', 'Annual'); `leaveStatus` is
    the approval state ('Pending'), not a type. Turn that into the wording the
    Planned Leaves sub-tasks use: 'Casual' -> 'Casual Leave'.
    """
    if rec.get("shortLeave") and not rec.get("leave"):
        return "Short Leave"
    text = str(rec.get("remarks") or "").strip()
    if not text:
        return "Planned Leave"
    if "leave" not in text.lower() and "holiday" not in text.lower():
        text = f"{text} Leave"
    return text.title()


def attendance_from_records(records: list) -> dict:
    """
    Turn the API's day records into the same shape the .xlsx path produces:
        {iso_date: {"minutes": int|None, "category": str, "leave": str|None}}

    The portal states each day outright — holiday / leave / dayOff / absent —
    so we trust those flags instead of reading words out of a row.
    """
    out: dict = {}
    for rec in records or []:
        date = rec.get("date")
        if not date:
            continue
        hours = rec.get("totalHours")
        if hours in (None, ""):
            hours = rec.get("workHours")
        minutes = int(round(float(hours) * 60)) if hours else None

        if rec.get("holiday"):
            out[date] = {"minutes": minutes, "category": "leave",
                         "leave": "Public Holiday"}
        elif rec.get("leave") and not minutes:
            out[date] = {"minutes": minutes, "category": "leave",
                         "leave": _leave_label(rec)}
        elif minutes:
            # Worked — including a day with a short leave on it, because the
            # short-leave top-up is worked out from the hours themselves.
            out[date] = {"minutes": minutes, "category": "work", "leave": None}
        elif rec.get("dayOff"):
            out[date] = {"minutes": None, "category": "weekend", "leave": None}
        elif rec.get("leave") or rec.get("shortLeave"):
            out[date] = {"minutes": minutes, "category": "leave",
                         "leave": _leave_label(rec)}
        elif rec.get("absent"):
            # Marked absent: the portal has no hours for it, and it is
            # certainly not leave. It is kept so the app can offer it - you may
            # well have worked it and forgotten to mark it - but nothing is
            # created for it unless you say so, and say how long.
            out[date] = {"minutes": None, "category": "absent", "leave": None}
        # A day still in progress is left out: there is nothing to log yet.
    return out


def fetch_attendance(headers: dict, start: str, end: str) -> dict:
    payload = _get_json(ATTENDANCE_URL, headers,
                        {"startDate": start, "endDate": end})
    records = (payload.get("data") or {}).get("attendanceRecords") or []
    return attendance_from_records(records)


# --------------------------------------------------------------------------- #
#  Getting a session out of the portal
# --------------------------------------------------------------------------- #
def _token_is_fresh(token: str) -> bool:
    """A token with at least a minute left on it. Unreadable ones are trusted."""
    exp = _jwt_expiry(token)
    return (exp - time.time() > 60) if exp else True


def _credentials(driver) -> tuple:
    """
    The bearer token + session id the SPA keeps in localStorage.

    The portal leaves the last token behind when it expires, so it is there to
    be read the instant the page opens — using it means an immediate 401. Only
    a token with life left in it counts; a dead one is cleared out so the app
    signs in again instead of handing us the same corpse next poll.
    """
    try:
        token = driver.execute_script("return localStorage.getItem('token')")
        if token and not _token_is_fresh(token):
            driver.execute_script(
                "localStorage.removeItem('token'); localStorage.removeItem('sessionId');")
            return None, "stale"
        return token, driver.execute_script("return localStorage.getItem('sessionId')")
    except Exception:  # noqa: BLE001
        return None, None


# The portal's bearer token lasts about half an hour, so hold on to it: a
# second fetch then costs two HTTPS calls instead of a whole browser.
_TOKEN: dict = {"headers": None, "expires": 0.0}


def _jwt_expiry(token: str) -> float:
    """The 'exp' claim as an epoch time, or 0 when it can't be read."""
    try:
        payload = token.split(".")[1]
        payload += "=" * (-len(payload) % 4)
        return float(json.loads(base64.urlsafe_b64decode(payload)).get("exp") or 0)
    except Exception:  # noqa: BLE001 — an opaque token is fine, just not cacheable
        return 0.0


def _remember(headers: dict, token: str) -> dict:
    exp = _jwt_expiry(token)
    lifetime = (exp - time.time()) if exp else 900
    _TOKEN.update(headers=headers, expires=time.time() + max(0.0, lifetime) - 60)
    return headers


def _cached_headers() -> Optional[dict]:
    if _TOKEN["headers"] and time.time() < _TOKEN["expires"]:
        return _TOKEN["headers"]
    return None


def forget_token() -> None:
    _TOKEN.update(headers=None, expires=0.0)


# The portal's sign-in control is a custom element (<i2c-button>), so it is
# found by its text rather than by tag, and clicked from script.
JS_CLICK_MICROSOFT = r"""
var hits = Array.from(document.querySelectorAll('*')).filter(function(el){
  var t = (el.innerText || el.value || '').trim();
  return /sign\s*in\s*with\s*microsoft/i.test(t) && t.length < 80;
});
var deep = hits.filter(function(el){          // the control, not its wrappers
  return !hits.some(function(o){ return o !== el && el.contains(o); });
});
var el = deep[0] || hits[hits.length - 1];
if (!el) return false;
el.click();
return true;
"""

MS_WAIT = 25       # seconds to let Microsoft bounce us back before giving up


def _click_microsoft(driver) -> bool:
    """Press the portal's 'Sign in with Microsoft'. True if it was there."""
    try:
        return bool(driver.execute_script(JS_CLICK_MICROSOFT))
    except Exception:  # noqa: BLE001
        return False


# --------------------------------------------------------------------------- #
#  The fetch job
# --------------------------------------------------------------------------- #
class AttendanceFetch:
    """Runs one portal fetch in the background and reports its progress."""

    def __init__(self) -> None:
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None
        self._cancel = threading.Event()
        self._launched = False
        self._state = {"state": "idle", "message": "", "days": 0,
                       "work_days": 0, "leave_days": 0, "note": "",
                       "range_start": "", "range_end": "",
                       "can_open_window": False}

    def _set(self, state: str, message: str, **extra) -> None:
        with self._lock:
            self._state = {"state": state, "message": message, "days": 0,
                           "work_days": 0, "leave_days": 0, "note": "",
                           "range_start": "", "range_end": "",
                           "can_open_window": False, **extra}

    def status(self) -> dict:
        with self._lock:
            state = dict(self._state)
        # A worker that died without reporting must never leave the page
        # spinning: if nothing is running, 'working' is a lie.
        if state["state"] == "working" and self._launched and not self.busy():
            state.update(state="error", can_open_window=True,
                         message="The fetch stopped unexpectedly. Try again.")
            with self._lock:
                self._state = dict(state)
        return state

    def busy(self) -> bool:
        return bool(self._thread and self._thread.is_alive())

    def cancel(self) -> None:
        self._cancel.set()
        self._set("idle", "")

    def start(self, release: str, start: str, end: str,
              on_success: Callable[[dict], None],
              interactive: bool = False) -> tuple:
        if self.busy():
            if not self._cancel.is_set():
                return True, ""
            self._thread.join(timeout=5)
            if self.busy():
                return False, "Still closing the previous fetch — try again."
        self._cancel.clear()
        self._launched = False
        self._set("working", "Starting…")
        self._thread = threading.Thread(
            target=self._run, args=(release, start, end, on_success, interactive),
            daemon=True)
        self._thread.start()
        self._launched = True
        return True, ""

    # ---- the actual work -------------------------------------------------- #
    def _session(self, interactive: bool):
        """Open the portal and lift its token. Raises with advice on failure."""
        cached = _cached_headers()
        if cached and not interactive:
            return cached

        # 'none' hands control back the instant navigation starts — this is a
        # single-page app, so waiting for a load event buys nothing.
        driver = sso_login.launch_browser(headless=not interactive,
                                          page_load_strategy="none")
        try:
            try:
                driver.set_page_load_timeout(sso_login.PAGE_LOAD_TIMEOUT)
                driver.get(PORTAL_URL)
            except Exception:  # noqa: BLE001 — it keeps loading underneath us
                pass

            deadline = time.time() + (WINDOW_TIMEOUT if interactive else TOKEN_TIMEOUT)
            login_since = clicked_at = last_click = None
            reloaded = False
            while time.time() < deadline:
                if self._cancel.is_set():
                    return None
                token, session_id = _credentials(driver)
                if token:
                    return _remember(_headers(token, session_id), token)
                if session_id == "stale" and not reloaded:
                    # Cleared an expired token: restart the app so it goes to
                    # its sign-in page rather than sitting on a dead session.
                    reloaded = True
                    self._set("working", "Session expired — signing in again…")
                    try:
                        driver.get(PORTAL_URL)
                    except Exception:  # noqa: BLE001
                        pass
                    login_since = clicked_at = last_click = None
                    continue

                try:
                    url = driver.current_url or ""
                except Exception:  # noqa: BLE001 — the browser went away
                    raise RuntimeError("The browser stopped before the portal "
                                       "signed in.")
                now = time.time()

                if "login.microsoftonline.com" in url:
                    # Bouncing through Microsoft. If it settles here it wants a
                    # person, and nothing in the background can be one.
                    self._set("working", "Signing in with Microsoft…")
                    if clicked_at and not interactive and now - clicked_at > MS_WAIT:
                        raise RuntimeError(_NOT_SIGNED_IN)
                elif "/login" in url:
                    # The portal's own page. The button is drawn by the app, so
                    # it may not exist for a second or two after the URL does —
                    # keep trying rather than giving up on the first look.
                    login_since = login_since or now
                    if last_click is None or now - last_click > CLICK_COOLDOWN:
                        if _click_microsoft(driver):
                            last_click = now
                            clicked_at = clicked_at or now
                            self._set("working", "Signing in with Microsoft…")
                        else:
                            last_click = now
                            self._set("working", "Opening the attendance portal…")
                    if (not interactive and clicked_at is None
                            and now - login_since > LOGIN_GRACE):
                        raise RuntimeError(_NOT_SIGNED_IN)
                else:
                    self._set("working", "Waiting for the portal… "
                                         f"({sso_login._short(url)})")
                time.sleep(POLL)
            raise RuntimeError(_NOT_SIGNED_IN)
        finally:
            try:
                driver.quit()
            except Exception:  # noqa: BLE001
                pass

    def _session_locked(self, interactive: bool):
        """A session, reusing the cached token when it is still good."""
        cached = _cached_headers()
        if cached and not interactive:
            return cached                      # no browser, no lock, no wait
        if not sso_login.BROWSER_LOCK.acquire(timeout=180):
            raise RuntimeError("The browser is busy with another step.")
        try:
            return self._session(interactive)
        finally:
            sso_login.BROWSER_LOCK.release()

    def _fetch(self, headers, release, start, end, on_success) -> None:
        note = ""
        # The dates the UI should offer: the release as the portal shows it,
        # which is a narrower span than the one we fetch (that runs on to the
        # exception cut-off so no worked day is missed).
        span_start = start
        span_end = end
        if release:
            self._set("working", f"Looking up release {release}…")
            found = release_range(headers, release)
            start, end = found["start"], found["end"]
            span_start, span_end = found["start"], found["release_end"]
            note = f"release {found['name']} ({span_start} → {span_end})"
        if not (start and end):
            return self._set("error", "No release and no dates to fetch for.")

        self._set("working", f"Reading your attendance for {start} → {end}…")
        attendance = fetch_attendance(headers, start, end)
        if not attendance:
            return self._set("error", "The portal returned no attendance for "
                                      f"{start} → {end}.")
        on_success(attendance)
        work = sum(1 for r in attendance.values() if r["category"] == "work")
        leaves = sum(1 for r in attendance.values() if r["category"] == "leave")
        self._set("done", f"Read {len(attendance)} day(s) from the portal.",
                  days=len(attendance), work_days=work, leave_days=leaves,
                  note=note, range_start=span_start or "", range_end=span_end or "")

    def _run(self, release, start, end, on_success, interactive=False) -> None:
        for attempt in (1, 2):
            try:
                self._set("working", "Opening the attendance portal…")
                headers = self._session_locked(interactive)
            except Exception as exc:  # noqa: BLE001
                return self._set("error", str(exc), can_open_window=True)
            if headers is None:
                return self._set("idle", "")          # cancelled
            try:
                return self._fetch(headers, release, start, end, on_success)
            except Exception as exc:  # noqa: BLE001
                # A cached token that expired early: drop it and mint one more.
                if "401" in str(exc) and attempt == 1:
                    forget_token()
                    continue
                return self._set("error", str(exc))


MANAGER = AttendanceFetch()
