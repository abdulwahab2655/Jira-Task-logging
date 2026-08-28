#!/usr/bin/env python3
"""
jira_logging_utility.py — Jira sub-task + timesheet automation.

Both the command line tool and the engine behind the web UI (app.py):
the hour rules, the grouping and the Jira client all live here.

WORKFLOW
--------
1. Log in to Jira (username / password, or PAT).
2. Ask which parent issue to create sub-tasks under (e.g. CR-10835).
3. Ask for a start date and an end date.
4. For every working day in that range it will:
       a) create an 8-hour sub-task under the parent
       b) log 8h of work against it, dated to that day
       c) transition it to "Done"
5. Repeat for each date, printing a running report.

By default only weekdays (Mon-Fri) get a sub-task. Pass --include-weekends
to log all 7 days.

USAGE
-----
    pip install requests openpyxl
    python jira_logging_utility.py             # fully interactive
    python jira_logging_utility.py --dry-run   # show what it WOULD do, no changes

You can also pre-fill answers with flags to skip the prompts:
    python jira_logging_utility.py \
        --base-url https://jira.company.com \
        --parent CR-10835 \
        --start 2026-08-01 --end 2026-08-15 \
        --summary "Development"

NOTES ON LOGIN
--------------
* Jira Server / Data Center: username + password often works, but many
  instances now require a Personal Access Token instead of a password.
  If password login is rejected, create a PAT in your Jira profile and
  paste that when asked for the password (leave username blank).
* Jira Cloud: use your email as the username and an API token as the
  password (basic password login is disabled on Cloud).
"""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import getpass
import os
import re
import sys
from dataclasses import dataclass
from typing import Optional

import requests
from concurrent.futures import ThreadPoolExecutor
from requests.adapters import HTTPAdapter
from requests.auth import HTTPBasicAuth


WORK_HOURS = "8h"          # time logged per sub-task
POOL_SIZE = 32             # sockets kept open per host; a run needs many at once


def _retry_policy():
    """
    Retry only what is safe to repeat - reads - and only on a blip or a rate
    limit. A create or a worklog is never retried here: the first attempt may
    well have landed, and a duplicate work log is worse than an error.
    """
    try:
        from urllib3.util.retry import Retry  # noqa: WPS433
    except Exception:  # noqa: BLE001 - an old urllib3 just means no retries
        return 0
    kwargs = {"total": 2, "backoff_factor": 0.4, "raise_on_status": False,
              "status_forcelist": (429, 500, 502, 503, 504)}
    try:
        return Retry(allowed_methods=frozenset({"GET"}), **kwargs)
    except TypeError:                       # urllib3 < 1.26 spelt it differently
        return Retry(method_whitelist=frozenset({"GET"}), **kwargs)


def in_parallel(fn, items, limit: int = 8) -> list:
    """
    Map `fn` over `items` at the same time, results in the order asked.

    Everything this is used for is either a read or a write against a
    different issue, so there is nothing to serialise - and the wait is all
    network.
    """
    items = list(items)
    if len(items) < 2:
        return [fn(item) for item in items]
    with ThreadPoolExecutor(max_workers=min(limit, len(items))) as pool:
        return list(pool.map(fn, items))


class LoginError(Exception):
    """Raised when authentication fails; message is already user-friendly."""


def _norm_sprint(name) -> str:
    """Normalise a sprint name so 'ST12-26.8', 'ST12 26.8' and 'st12_26.8' match."""
    return re.sub(r"[\s._-]+", "", str(name).lower())


# Sprint names are written differently by every team:
#   'ST12-26.8'   'ST-19 - 26.08 Release'   'ST 7 26.8 Sprint'   'ST19_26.08'
# Rather than pattern-match the whole string, pull the two things that
# actually identify a sprint out of it: the team number and the release number.
_TEAM_RE = re.compile(r"\bst[\s._-]*0*(\d+)", re.I)
_RELEASE_RE = re.compile(r"(\d{1,4})\s*\.\s*0*(\d{1,3})(?:\s*\.\s*0*(\d{1,3}))?")


def parse_team_number(text) -> Optional[int]:
    """'ST12' / 'st-19' / 'ST 7' / '12' -> 12, 19, 7, 12.  None if absent."""
    s = str(text or "").strip()
    if not s:
        return None
    m = _TEAM_RE.search(s)
    if m:
        return int(m.group(1))
    m = re.fullmatch(r"0*(\d+)", s)          # a bare number is a team number too
    return int(m.group(1)) if m else None


def parse_release_number(text) -> Optional[tuple]:
    """
    '26.8' / '26.08' / '26.08 Release' -> (26, 8, None); '26.8.1' -> (26, 8, 1).
    A bare '26' -> (26, None, None), meaning "any release in 26".
    """
    s = str(text or "").strip()
    if not s:
        return None
    m = _RELEASE_RE.search(s)
    if m:
        return (int(m.group(1)), int(m.group(2)),
                int(m.group(3)) if m.group(3) else None)
    m = re.fullmatch(r"0*(\d+)", s)
    return (int(m.group(1)), None, None) if m else None


def sprint_sequence(name: str) -> int:
    """Trailing sprint number, as in '26.08 - ST4 Sprint 8' -> 8. -1 if none."""
    m = re.search(r"\bsprint[\s._#-]*0*(\d+)", str(name), re.I)
    return int(m.group(1)) if m else -1


def sprint_name_matches(name: str, team: int, release: tuple) -> bool:
    """True when a sprint name belongs to this team AND this release."""
    if parse_team_number(name) != team:
        return False
    found = parse_release_number(name)
    if not found:
        return False
    if found[0] != release[0]:
        return False
    # A part the caller left out ('26' or '26.8') matches anything in that slot,
    # so '26.8' finds '26.08' and '26.8.1' alike.
    for want, got in zip(release[1:], found[1:]):
        if want is not None and want != got:
            return False
    return True


def _clean_server_text(text: str, limit: int = 160) -> str:
    """Strip HTML/whitespace from a server error body so it reads cleanly."""
    if not text:
        return ""
    no_tags = re.sub(r"<[^>]+>", " ", text)          # drop HTML tags
    collapsed = re.sub(r"\s+", " ", no_tags).strip()
    return collapsed[:limit]


# --------------------------------------------------------------------------- #
#  Result tracking
# --------------------------------------------------------------------------- #
@dataclass
class DayResult:
    date: str
    created_key: Optional[str] = None
    logged: bool = False
    done: bool = False
    error: Optional[str] = None
    hours: str = ""            # what the sub-task holds in total
    days: int = 1              # how many work logs went into it

    @property
    def ok(self) -> bool:
        return bool(self.created_key and self.logged and self.done and not self.error)

    def __str__(self) -> str:
        if self.error:
            return f"[FAIL] {self.date}  {self.created_key or '(not created)'}  -> {self.error}"
        logged = (f"logged {self.hours}" if self.hours else "-")
        if self.logged and self.days > 1:
            logged += f" over {self.days} days"
        marks = [
            "created" if self.created_key else "-",
            logged if self.logged else "-",
            "Done" if self.done else "-",
        ]
        return f"[ OK ] {self.date}  {self.created_key:<14} {' | '.join(marks)}"


# --------------------------------------------------------------------------- #
#  Jira client
# --------------------------------------------------------------------------- #
class JiraClient:
    def __init__(self, base_url: str, username: str = "", password: str = "",
                 api_version: str = "2", timeout: int = 30, verify_ssl: bool = True,
                 session_cookie: str = ""):
        self.base_url = base_url.rstrip("/")
        self.api = f"{self.base_url}/rest/api/{api_version}"
        self.timeout = timeout
        self.session = requests.Session()
        self.session.verify = verify_ssl
        # Without this, requests keeps ten sockets per host and quietly drops
        # the rest - so a parallel run pays for a fresh TLS handshake per call.
        adapter = HTTPAdapter(pool_connections=POOL_SIZE, pool_maxsize=POOL_SIZE,
                              max_retries=_retry_policy())
        self.session.mount("https://", adapter)
        self.session.mount("http://", adapter)
        self.session.headers.update({"Accept": "application/json",
                                     "Content-Type": "application/json"})
        self.assignee_ref: Optional[dict] = None  # set by verify_login()
        # (project, status, Task) -> (transition id, screen payload). A Jira
        # workflow is the same for every sub-task in a project, so the walk to
        # Done can be learnt from the first one and reused by the rest.
        self._steps: dict = {}
        self.me: dict = {}                        # set by verify_login()
        self.tz = None                            # set by verify_login()

        if session_cookie:
            # Microsoft SSO path: reuse the browser's authenticated Jira cookie.
            cookie = session_cookie.strip()
            # Accept a bare JSESSIONID value or a full "name=value; ..." string.
            if "=" not in cookie:
                cookie = f"JSESSIONID={cookie}"
            self.session.headers["Cookie"] = cookie
        elif username:
            self.session.auth = HTTPBasicAuth(username, password)
        else:
            # Blank username -> treat the password as a bearer PAT.
            self.session.headers["Authorization"] = f"Bearer {password}"

    # ---- login check ----------------------------------------------------- #
    def verify_login(self) -> str:
        """Return the display name of the logged-in user, and cache identity.

        Raises LoginError with a clear, already-formatted message on failure.
        """
        try:
            r = self.session.get(f"{self.base_url}/rest/api/2/myself",
                                 timeout=self.timeout)
        except requests.exceptions.SSLError:
            raise LoginError(
                "SSL certificate check failed. If your Jira uses an internal "
                "certificate, re-run with --no-verify-ssl."
            )
        except requests.exceptions.ConnectionError:
            raise LoginError(
                f"Couldn't reach {self.base_url}. Check the URL and your network/VPN."
            )
        except requests.exceptions.Timeout:
            raise LoginError(f"Connection to {self.base_url} timed out. Try again.")

        if r.status_code == 200:
            data = r.json()
            if data.get("accountId"):
                self.assignee_ref = {"accountId": data["accountId"]}
            elif data.get("name"):
                self.assignee_ref = {"name": data["name"]}
            else:
                self.assignee_ref = None
            # Keep the whole identity, not just the assignee ref: reading back
            # what is already logged means telling your worklogs from everyone
            # else's on the same issue.
            self.me = {"accountId": data.get("accountId") or "",
                       "key": data.get("key") or "",
                       "name": data.get("name") or "",
                       "displayName": data.get("displayName") or "",
                       "timeZone": data.get("timeZone") or ""}
            self.tz = _zone(self.me["timeZone"])
            return data.get("displayName") or data.get("name") or "user"

        # --- friendly messages for the common failures ---
        captcha = r.headers.get("X-Authentication-Denied-Reason", "")
        if "CAPTCHA" in captcha.upper():
            raise LoginError(
                "Jira is asking for a CAPTCHA after too many failed logins. "
                "Open Jira in your browser, sign in once (solving the CAPTCHA), "
                "then try again. A Personal Access Token avoids this."
            )
        if r.status_code == 401:
            raise LoginError(
                "Wrong username or password (HTTP 401). Note: many i2c/Server "
                "instances block password login — create a Personal Access Token "
                "in your Jira profile, leave the username blank, and paste the "
                "token as the password."
            )
        if r.status_code == 403:
            raise LoginError(
                "Access denied (HTTP 403). Your account may be locked or lacks "
                "API permission. Sign in via the browser first, or use a PAT."
            )
        raise LoginError(
            f"Login failed (HTTP {r.status_code}). {_clean_server_text(r.text)}".strip()
        )

    # ---- parent / project lookup ---------------------------------------- #
    def get_parent_context(self, parent_key: str) -> dict:
        """Fetch project key + a reasonable summary base from the parent issue."""
        r = self.session.get(f"{self.api}/issue/{parent_key}",
                             params={"fields": "project,summary"}, timeout=self.timeout)
        r.raise_for_status()
        f = r.json()["fields"]
        return {"project_key": f["project"]["key"], "summary": f.get("summary", "")}

    def find_subtask_type_id(self) -> str:
        """Find the issue-type id whose 'subtask' flag is true."""
        r = self.session.get(f"{self.api}/issuetype", timeout=self.timeout)
        r.raise_for_status()
        for it in r.json():
            if it.get("subtask"):
                return it["id"]
        raise RuntimeError("No sub-task issue type found on this Jira instance.")

    def find_sprint_id(self, board_id: int, sprint_name: str) -> tuple:
        """
        Look up a sprint by name on a board (via the Agile API).
        Matches flexibly so 'ST12-26.8' also matches 'ST12 26.8', etc.
        Returns (sprint_id, actual_name) or (None, None).
        """
        target = _norm_sprint(sprint_name)
        for sp in self.board_sprints(board_id):
            if _norm_sprint(sp["name"]) == target:
                return sp["id"], sp["name"]
        return None, None

    def find_sprint_anywhere(self, sprint_name: str) -> Optional[dict]:
        """
        Find a sprint by name without knowing its board, using Jira's own
        sprint picker (the same search the sprint field uses in the UI).
        Returns {'id', 'name'} or None.
        """
        target = _norm_sprint(sprint_name)
        try:
            r = self.session.get(
                f"{self.base_url}/rest/greenhopper/1.0/sprint/picker",
                params={"query": sprint_name}, timeout=self.timeout,
            )
        except requests.RequestException:
            return None
        if r.status_code != 200:
            return None
        try:
            data = r.json()
        except ValueError:
            return None
        # 'suggestions' holds the open sprints, 'allMatches' includes closed ones.
        for bucket in ("suggestions", "allMatches"):
            for sp in data.get(bucket) or []:
                if _norm_sprint(sp.get("name", "")) == target and sp.get("id"):
                    return {"id": sp["id"], "name": sp.get("name") or sprint_name}
        return None

    def _picker_query(self, query: str) -> list[dict]:
        """Ask Jira's sprint picker for sprints whose name contains `query`."""
        try:
            r = self.session.get(
                f"{self.base_url}/rest/greenhopper/1.0/sprint/picker",
                params={"query": query}, timeout=self.timeout,
            )
        except requests.RequestException:
            return []
        if r.status_code != 200:
            return []
        try:
            data = r.json()
        except ValueError:
            return []
        out, seen = [], set()
        # 'suggestions' are the open sprints; 'allMatches' also includes closed.
        for bucket, default_state in (("suggestions", "open"), ("allMatches", "")):
            for sp in data.get(bucket) or []:
                sid, name = sp.get("id"), sp.get("name")
                if not sid or not name or sid in seen:
                    continue
                seen.add(sid)
                out.append({"id": sid, "name": name,
                            "state": (sp.get("stateKey") or default_state).lower(),
                            "board_id": sp.get("boardId")})
        return out

    def board_sprints(self, board_id) -> list[dict]:
        """Every sprint on a board, as [{'id','name','state','board_id'}, ...]."""
        out, start = [], 0
        while True:
            r = self.session.get(
                f"{self.base_url}/rest/agile/1.0/board/{board_id}/sprint",
                params={"startAt": start, "maxResults": 50,
                        "state": "active,future,closed"},
                timeout=self.timeout,
            )
            if r.status_code != 200:
                raise RuntimeError(
                    f"Couldn't list sprints on board {board_id} "
                    f"(HTTP {r.status_code}): {_clean_server_text(r.text)}"
                )
            data = r.json()
            for sp in data.get("values", []):
                out.append({"id": sp.get("id"), "name": sp.get("name", ""),
                            "state": str(sp.get("state", "")).lower(),
                            "board_id": sp.get("originBoardId") or board_id})
            if data.get("isLast", True) or not data.get("values"):
                break
            start += len(data["values"])
        return out

    def find_boards(self, name: str, limit: int = 50) -> list[dict]:
        """Boards whose name contains `name`. [] when the search isn't available."""
        r = self.session.get(f"{self.base_url}/rest/agile/1.0/board",
                             params={"name": name, "maxResults": limit},
                             timeout=self.timeout)
        if r.status_code != 200:
            return []
        try:
            return r.json().get("values", []) or []
        except ValueError:
            return []

    def find_release_sprint(self, st: str, release: str) -> dict:
        """
        Find the sprint for a team + release however it happens to be named —
        'ST12-26.8', 'ST-19 - 26.08 Release', 'ST 7 26.08' all resolve the same.

        Returns {'id', 'name', 'board_id', 'state', 'others': [names]}.
        Raises RuntimeError naming the near misses when nothing matches.
        """
        team = parse_team_number(st)
        release_no = parse_release_number(release)
        if team is None:
            raise RuntimeError(f"Couldn't read a team number from '{st}'. "
                               "Use something like ST12 or ST-19.")
        if release_no is None:
            raise RuntimeError(f"Couldn't read a release number from '{release}'. "
                               "Use something like 26.8 or 26.08.")

        major, minor = release_no[0], release_no[1]
        # Search by release first (few hits instance-wide), then by team.
        queries = [f"{major}.{minor}", f"{major}.{minor:02d}"] if minor is not None else []
        queries += [f"ST{team}", f"ST-{team}", f"ST {team}", str(major)]

        seen: dict = {}
        matches: dict = {}
        near: list[str] = []

        def sift(entries) -> None:
            for e in entries:
                sid, name = e.get("id"), e.get("name") or ""
                if not sid or sid in seen:
                    continue
                seen[sid] = e
                if sprint_name_matches(name, team, release_no):
                    matches[sid] = e
                elif parse_team_number(name) == team:
                    near.append(name)

        # The picker calls are independent reads: they go out together, and
        # are sifted in the order asked, so the answer is unchanged.
        wanted = list(dict.fromkeys(queries))
        for entries in in_parallel(self._picker_query, wanted, limit=6):
            sift(entries)
            if matches:
                break

        # No picker (or it found nothing): go via the team's boards instead.
        if not matches:
            boards, board_seen = [], set()
            for q in (f"ST{team}", f"ST-{team}", f"ST {team}"):
                for b in self.find_boards(q):
                    if b.get("id") and b["id"] not in board_seen:
                        board_seen.add(b["id"])
                        boards.append(b)
            for b in boards[:10]:             # keep the scan bounded
                try:
                    sift(self.board_sprints(b["id"]))
                except RuntimeError:
                    continue
                if matches:
                    break

        if not matches:
            detail = ""
            if near:
                shown = ", ".join(sorted(dict.fromkeys(near))[:5])
                detail = f" Sprints found for ST{team}: {shown}."
            raise RuntimeError(
                f"No sprint matched ST{team} release {release}.{detail}"
            )

        # One team can run several sprints in a release ('… Sprint 8', '… Sprint 9'),
        # so prefer a live one, then the latest sprint number, then the plainest name.
        order = {"active": 0, "open": 1, "future": 2, "": 3, "closed": 4}
        ranked = sorted(matches.values(),
                        key=lambda e: (order.get(e.get("state", ""), 3),
                                       -sprint_sequence(e["name"]), len(e["name"])))
        best = ranked[0]
        board_id = best.get("board_id")
        if not board_id:
            board_id = self.get_sprint(best["id"]).get("originBoardId")
        return {"id": best["id"], "name": best["name"], "board_id": board_id,
                "state": best.get("state", ""),
                "others": [e["name"] for e in ranked[1:]]}

    def get_sprint(self, sprint_id) -> dict:
        """Fetch a sprint's details (name, state, originBoardId). {} if unavailable."""
        r = self.session.get(f"{self.base_url}/rest/agile/1.0/sprint/{sprint_id}",
                             timeout=self.timeout)
        if r.status_code != 200:
            return {}
        try:
            return r.json()
        except ValueError:
            return {}

    def resolve_sprint(self, sprint: str, board_id: Optional[int] = None) -> Optional[dict]:
        """
        Locate a sprint and the board it came from. A board id is only a hint:
        if it's missing — or the sprint isn't on it — we look the sprint up by
        name across the whole instance and read its board back off the sprint.
        Returns {'id', 'name', 'board_id'} or None.
        """
        if board_id:
            try:
                sid, actual = self.find_sprint_id(board_id, sprint)
            except RuntimeError:
                sid = actual = None          # bad/inaccessible board -> auto-detect
            if sid is not None:
                return {"id": sid, "name": actual or sprint, "board_id": int(board_id)}

        hit = self.find_sprint_anywhere(sprint)
        if not hit:
            return None
        details = self.get_sprint(hit["id"])
        return {"id": hit["id"],
                "name": details.get("name") or hit["name"],
                "board_id": details.get("originBoardId")}

    def issues_in_sprint(self, sprint_id) -> list[tuple]:
        """
        [(key, summary, project_key), ...] for the non-sub-task issues.

        The project comes back with them because the create metadata - the
        slowest call in the whole app - is per project, not per ticket: one
        lookup then serves every ticket in it.
        """
        issues, start = [], 0
        while True:
            r = self.session.get(
                f"{self.base_url}/rest/agile/1.0/sprint/{sprint_id}/issue",
                params={"fields": "summary,issuetype,project", "startAt": start,
                        "maxResults": 100},
                timeout=self.timeout,
            )
            if r.status_code != 200:
                raise RuntimeError(
                    f"Couldn't read sprint {sprint_id} issues "
                    f"(HTTP {r.status_code}): {_clean_server_text(r.text)}"
                )
            data = r.json()
            for it in data.get("issues", []):
                f = it.get("fields", {})
                if f.get("issuetype", {}).get("subtask"):
                    continue
                issues.append((it["key"], f.get("summary", ""),
                               (f.get("project") or {}).get("key", "")))
            total = data.get("total", 0)
            start += len(data.get("issues", []))
            if start >= total or not data.get("issues"):
                break
        return issues

    def _issues_by_jql(self, sprint: str) -> list[tuple]:
        """Last-resort lookup by sprint name through JQL. [] if it doesn't work."""
        jql = (f'sprint = "{_jql_text(sprint)}" AND '
               'issuetype not in subTaskIssueTypes() ORDER BY key ASC')
        r = self.session.get(
            f"{self.api}/search",
            params={"jql": jql, "fields": "summary,project", "maxResults": 100},
            timeout=self.timeout,
        )
        if r.status_code != 200:
            return []
        return [(it["key"], it.get("fields", {}).get("summary", ""),
                 ((it.get("fields", {}).get("project") or {}).get("key", "")))
                for it in r.json().get("issues", [])]

    def search_sprint(self, sprint: str, board_id: Optional[int] = None) -> tuple:
        """
        Return (issues, meta) for a sprint, where meta carries what we resolved:
        {'sprint_id', 'sprint_name', 'board_id'}. The board id is optional —
        pass one to pin the search, or leave it out and it gets detected.
        """
        meta = {"sprint_id": None, "sprint_name": sprint, "board_id": board_id}
        found = self.resolve_sprint(sprint, board_id)
        if found:
            meta.update(sprint_id=found["id"], sprint_name=found["name"],
                        board_id=found["board_id"])
            return self.issues_in_sprint(found["id"]), meta

        issues = self._issues_by_jql(sprint)
        if issues:
            return issues, meta
        raise RuntimeError(
            f"Sprint '{sprint}' not found"
            + (f" on board {board_id}, and no sprint of that name exists elsewhere. "
               if board_id else ". ")
            + "Check the spelling, or pass the board id explicitly."
        )

    def get_project_subtask_context(self, project_key: str) -> tuple:
        """
        For a project, return (subtask_type_id, create_fields) for its sub-task
        issue type, reading the project's create metadata directly.
        """
        r = self.session.get(
            f"{self.api}/issue/createmeta",
            params={"projectKeys": project_key,
                    "expand": "projects.issuetypes.fields"},
            timeout=self.timeout,
        )
        if r.status_code == 200:
            for proj in r.json().get("projects", []):
                if proj.get("key") != project_key:
                    continue
                for it in proj.get("issuetypes", []):
                    if it.get("subtask"):
                        return it.get("id"), (it.get("fields") or {})
        # Fallback: global sub-task type id, fields fetched separately.
        stid = self.find_subtask_type_id()
        return stid, self.get_create_fields(project_key, stid)

    def get_create_fields(self, project_key: str, issuetype_id: str) -> dict:
        """
        Return the create-screen field metadata for this project + sub-task type,
        normalised to {field_id: {name, allowedValues, ...}}.
        Handles both the classic (Server/DC) and newer (Cloud) createmeta APIs.
        """
        # Classic Server/DC endpoint.
        r = self.session.get(
            f"{self.api}/issue/createmeta",
            params={"projectKeys": project_key, "issuetypeIds": issuetype_id,
                    "expand": "projects.issuetypes.fields"},
            timeout=self.timeout,
        )
        if r.status_code == 200:
            data = r.json()
            for proj in data.get("projects", []):
                for it in proj.get("issuetypes", []):
                    if str(it.get("id")) == str(issuetype_id):
                        return it.get("fields", {}) or {}

        # Newer Cloud endpoint (paginated list of field objects).
        r = self.session.get(
            f"{self.api}/issue/createmeta/{project_key}/issuetypes/{issuetype_id}",
            timeout=self.timeout,
        )
        if r.status_code == 200:
            out: dict = {}
            for f in r.json().get("values", []):
                fid = f.get("fieldId") or f.get("key")
                if fid:
                    out[fid] = f
            return out

        return {}

    def resolve_task_field(self, create_fields: dict, task_value: str):
        """
        Find the custom 'Task' dropdown field id and the option matching
        `task_value`. Returns (field_id, value_ref) — either may be None.
        """
        for fid, meta in create_fields.items():
            if str(meta.get("name", "")).strip().lower() == "task":
                for o in meta.get("allowedValues", []) or []:
                    if str(o.get("value", "")).strip().lower() == task_value.strip().lower():
                        return fid, {"id": str(o["id"])}
                return fid, None  # field exists but no matching option
        return None, None

    def get_task_options(self, create_fields: dict):
        """
        Return (field_id, [(value, id), ...]) for the 'Task' dropdown so the
        caller can present the real option list to the user. Skips the
        placeholder 'None' entry. Either part may be empty if not found.
        """
        for fid, meta in create_fields.items():
            if str(meta.get("name", "")).strip().lower() == "task":
                opts = [
                    (str(o.get("value", "")), str(o.get("id")))
                    for o in (meta.get("allowedValues") or [])
                    if str(o.get("value", "")).strip().lower() != "none"
                ]
                return fid, opts
        return None, []

    # ---- create --------------------------------------------------------- #
    def create_subtask(self, project_key: str, parent_key: str,
                       subtask_type_id: str, summary: str,
                       extra_fields: Optional[dict] = None) -> str:
        fields = {
            "project": {"key": project_key},
            "parent": {"key": parent_key},
            "summary": summary,
            "issuetype": {"id": subtask_type_id},
        }
        if extra_fields:
            fields.update(extra_fields)
        r = self.session.post(f"{self.api}/issue", json={"fields": fields},
                             timeout=self.timeout)
        if r.status_code in (200, 201):
            return r.json()["key"]
        raise RuntimeError(f"Create failed (HTTP {r.status_code}): {_clean_server_text(r.text)}")

    # ---- log work ------------------------------------------------------- #
    @staticmethod
    def _started(date_str: str) -> str:
        dt = _dt.datetime.strptime(date_str, "%Y-%m-%d").replace(hour=9).astimezone()
        return dt.strftime("%Y-%m-%dT%H:%M:%S.000%z")

    def log_work(self, issue_key: str, date_str: str,
                 time_spent: str = WORK_HOURS, comment: str = "") -> None:
        payload = {"timeSpent": time_spent, "started": self._started(date_str)}
        if comment:
            payload["comment"] = comment
        r = self.session.post(f"{self.api}/issue/{issue_key}/worklog",
                             json=payload, timeout=self.timeout)
        if r.status_code not in (200, 201):
            raise RuntimeError(f"Worklog failed (HTTP {r.status_code}): {_clean_server_text(r.text)}")

    # ---- read back what is already logged -------------------------------- #
    def _is_me(self, author: dict) -> bool:
        """True when a worklog was written by the signed-in user."""
        if not isinstance(author, dict):
            return False
        me = self.me or {}
        # Whichever id both sides carry is the one that decides: accountId on
        # Cloud, name/key on Server. Falling through to the display name is a
        # last resort for instances that return nothing else.
        for field in ("accountId", "key", "name"):
            mine, theirs = me.get(field), author.get(field)
            if mine and theirs:
                return str(mine).strip().lower() == str(theirs).strip().lower()
        mine = str(me.get("displayName") or "").strip().lower()
        return bool(mine) and mine == str(author.get("displayName") or "").strip().lower()

    @staticmethod
    def _parse_started(text: str):
        """A worklog's `started` string as an aware datetime, or None."""
        for fmt in ("%Y-%m-%dT%H:%M:%S.%f%z", "%Y-%m-%dT%H:%M:%S%z"):
            try:
                return _dt.datetime.strptime(str(text), fmt)
            except (ValueError, TypeError):
                continue
        return None

    def worklog_date(self, started) -> str:
        """
        The day a worklog belongs to, counted the way you see it in Jira.

        Jira answers with the *instance's* offset, not yours: 9am in Karachi
        comes back as `2026-08-10T21:00:00.000-0700`, so reading the first ten
        characters files a Wednesday's work under Tuesday. The stamp is moved
        into your own timezone first, which is what Jira's timesheet screens do
        and what makes these numbers match them.
        """
        text = str(started or "")
        parsed = self._parse_started(text)
        if parsed is None:
            return text[:10]
        # astimezone(None) means "this machine's zone", the fallback when the
        # profile's zone can't be resolved.
        return parsed.astimezone(self.tz).date().isoformat()

    def _issue_worklogs(self, issue_key: str) -> list[dict]:
        """Every worklog on an issue. The search response truncates at 20."""
        out, start = [], 0
        while True:
            r = self.session.get(f"{self.api}/issue/{issue_key}/worklog",
                                 params={"startAt": start, "maxResults": 1000},
                                 timeout=self.timeout)
            if r.status_code != 200:
                return out
            data = r.json()
            batch = data.get("worklogs") or []
            out.extend(batch)
            start += len(batch)
            if not batch or start >= int(data.get("total") or 0):
                return out

    def my_worklogs(self, start: str, end: str) -> dict:
        """
        What you have *already* logged in Jira between two dates (inclusive).

        Returns
            {iso_date: {"minutes": int,                 # your total that day
                        "parents": {parent_key: mins},  # by the issue's parent
                        "issues":  {issue_key: mins}}}

        Jira's `worklogDate` search finds the issues; every worklog on them is
        then checked one by one, because the same issue can carry other
        people's time and other days' time too.

        The search runs a day wide at each end: Jira matches `worklogDate` in
        its own timezone, which is not necessarily yours, so the edges of the
        range are settled by worklog_date() rather than by the search.
        """
        jql = (f'worklogAuthor = currentUser() AND '
               f'worklogDate >= "{shift_date(start, -1)}" AND '
               f'worklogDate <= "{shift_date(end, 1)}"')
        issues, at = [], 0
        while True:
            r = self.session.get(
                f"{self.api}/search",
                params={"jql": jql, "fields": "worklog,parent,summary",
                        "startAt": at, "maxResults": 50},
                timeout=self.timeout,
            )
            if r.status_code != 200:
                raise RuntimeError(
                    "Couldn't read what is already logged in Jira "
                    f"(HTTP {r.status_code}): {_clean_server_text(r.text)}"
                )
            data = r.json()
            batch = data.get("issues") or []
            issues.extend(batch)
            at += len(batch)
            if not batch or at >= int(data.get("total") or 0):
                break

        # An issue whose worklogs came back empty (the search did not expand
        # the field) or short of its own total (it truncated at 20) has to be
        # asked directly - and those asks are independent, so they go together.
        def whole(it) -> bool:
            container = (it.get("fields") or {}).get("worklog") or {}
            have = container.get("worklogs") or []
            return bool(have) and int(container.get("total") or 0) <= len(have)

        short = [it.get("key") or "" for it in issues if not whole(it)]
        fetched = dict(zip(short, in_parallel(self._issue_worklogs, short, limit=6)))

        out: dict = {}
        for it in issues:
            key = it.get("key") or ""
            f = it.get("fields") or {}
            parent = (f.get("parent") or {}).get("key") or ""
            logs = (f.get("worklog") or {}).get("worklogs") or []
            if key in fetched:
                logs = fetched[key]
            for w in logs:
                if not self._is_me(w.get("author") or {}):
                    continue
                date = self.worklog_date(w.get("started"))
                if not date or date < start or date > end:
                    continue
                mins = int(round(float(w.get("timeSpentSeconds") or 0) / 60))
                if not mins:
                    mins = duration_minutes(w.get("timeSpent"))   # e.g. '1h 30m'
                if mins <= 0:
                    continue
                day = out.setdefault(date, {"minutes": 0, "parents": {}, "issues": {}})
                day["minutes"] += mins
                day["issues"][key] = day["issues"].get(key, 0) + mins
                if parent:
                    day["parents"][parent] = day["parents"].get(parent, 0) + mins
        return out

    # ---- transition helpers --------------------------------------------- #
    def _status_and_transitions(self, issue_key: str) -> tuple:
        """
        Status *and* the transitions available from it, in one request.

        Asking separately costs two round trips per status step — four per
        sub-task — and on a 20-day run that is most of a minute spent waiting
        for answers the same call already contains.
        """
        r = self.session.get(f"{self.api}/issue/{issue_key}",
                             params={"fields": "status",
                                     "expand": "transitions.fields"},
                             timeout=self.timeout)
        r.raise_for_status()
        data = r.json()
        return (data["fields"]["status"]["name"], data.get("transitions") or [])

    def _get_transitions(self, issue_key: str) -> list[dict]:
        r = self.session.get(f"{self.api}/issue/{issue_key}/transitions",
                             params={"expand": "transitions.fields"},
                             timeout=self.timeout)
        r.raise_for_status()
        return r.json().get("transitions", [])

    @staticmethod
    def _find_transition(transitions: list[dict], target_name: str) -> Optional[dict]:
        target = target_name.strip().lower()
        for t in transitions:
            name = t.get("name", "").lower()
            to_status = t.get("to", {}).get("name", "").lower()
            if target in (name, to_status):
                return t
        return None

    def _build_screen_fields(self, transition: dict, task_value: str,
                             check_all: bool) -> dict:
        """
        Build the `fields` payload for a transition screen.

        * Checklist fields (array of options): tick ONLY the checklist that
          matches the selected Task. e.g. Task = "Development" -> tick every
          box in "Development Checklist" and leave all other checklists
          (Analysis, Testing, Design, ...) untouched.
        * Single-select fields (like 'Task') -> set to the value matching
          `task_value` if present, else fall back to the default / first option.
        """
        payload: dict = {}
        meta = transition.get("fields") or {}
        target_task = task_value.strip().lower()

        for fid, fmeta in meta.items():
            schema = fmeta.get("schema", {})
            ftype = schema.get("type")
            items = schema.get("items")
            allowed = fmeta.get("allowedValues", []) or []
            field_name = str(fmeta.get("name", "")).strip().lower()

            # Checklist / multi-select fields: only fill the one for this Task.
            if ftype == "array" and items in ("option", "string"):
                # Normalise "<Task> Checklist" -> "<task>" for comparison.
                name_key = field_name
                if name_key.endswith("checklist"):
                    name_key = name_key[: -len("checklist")].strip()
                is_matching_checklist = (name_key == target_task)

                if check_all and allowed and is_matching_checklist:
                    payload[fid] = [{"id": str(o["id"])} for o in allowed if "id" in o]
                # Any other checklist is intentionally left blank.
                continue

            # Single select (dropdown), e.g. the "Task" field.
            if ftype == "option":
                chosen = None
                for o in allowed:
                    if str(o.get("value", "")).strip().lower() == target_task:
                        chosen = {"id": str(o["id"])}
                        break
                if chosen is None and fmeta.get("hasDefaultValue"):
                    continue  # let Jira apply the default it already shows
                if chosen is None and allowed:
                    chosen = {"id": str(allowed[0]["id"])}
                if chosen is not None:
                    payload[fid] = chosen
                continue

            # Anything else that is required but we can't fill -> surface it.
            if fmeta.get("required") and not fmeta.get("hasDefaultValue"):
                raise RuntimeError(
                    f"Screen field '{fmeta.get('name', fid)}' is required and "
                    "can't be auto-filled. Tell me what value to use."
                )
        return payload

    def _do_transition(self, issue_key: str, transition_id: str, fields: dict) -> None:
        body: dict = {"transition": {"id": transition_id}}
        if fields:
            body["fields"] = fields
        r = self.session.post(f"{self.api}/issue/{issue_key}/transitions",
                             json=body, timeout=self.timeout)
        if r.status_code not in (200, 204):
            raise RuntimeError(f"Transition failed (HTTP {r.status_code}): {_clean_server_text(r.text)}")

    def move_through(self, issue_key: str, path: list[str], *,
                     task_value: str = "Development", check_all: bool = True,
                     project: str = "") -> None:
        """
        Walk the issue through each status in `path` in order
        (e.g. ['In Progress', 'Done']). Skips a step if the issue is already
        in that status; fills any screen (like the Done checklist) as it goes.

        The first sub-task of a project pays for reading its workflow; the rest
        reuse what it learnt, which saves two round trips each. If a reused
        transition is refused - the issue was not where we assumed - the slow,
        certain path runs instead, so the shortcut can only cost time, never
        correctness.
        """
        for step in path:
            known = self._steps.get((project, step, task_value))
            if known is not None:
                try:
                    self._do_transition(issue_key, known[0], known[1])
                    continue
                except RuntimeError:
                    pass                    # ask properly, below
            current, transitions = self._status_and_transitions(issue_key)
            if current.strip().lower() == step.strip().lower():
                continue  # already there
            match = self._find_transition(transitions, step)
            if match is not None and "fields" not in match:
                # Some instances only attach screen metadata to the dedicated
                # endpoint; fall back rather than skip a required field.
                match = self._find_transition(self._get_transitions(issue_key), step)
            if match is None:
                available = ", ".join(t.get("name", "?") for t in transitions) or "(none)"
                raise RuntimeError(
                    f"No '{step}' transition from status '{current}'. Options: {available}"
                )
            fields = self._build_screen_fields(match, task_value, check_all)
            self._do_transition(issue_key, match["id"], fields)
            self._steps[(project, step, task_value)] = (match["id"], fields)


# --------------------------------------------------------------------------- #
#  Date helpers
# --------------------------------------------------------------------------- #
def _zone(name: str):
    """
    A tzinfo for a Jira profile timezone ('Asia/Karachi'), or None for "use
    this machine's". Windows has no IANA database unless `tzdata` is installed,
    and the fallback is the right answer anyway: the machine sits in the same
    timezone as the person using it.
    """
    if not name:
        return None
    try:
        from zoneinfo import ZoneInfo  # noqa: WPS433
        return ZoneInfo(str(name))
    except Exception:  # noqa: BLE001 - no database, or a name it doesn't know
        return None


def _jql_text(value: str) -> str:
    """
    A value safe to sit inside a quoted JQL string.

    Sprint names are typed by hand, and one containing a quote used to end the
    string early: the query became malformed, Jira answered 400, and the app
    reported "no issues found" — the right answer to the wrong question. Dates
    reach JQL through date.fromisoformat(), so they need none of this.
    """
    return str(value or "").replace("\\", "\\\\").replace('"', '\\"')


def shift_date(iso: str, days: int) -> str:
    """'2026-08-10' shifted by whole days: -1 -> '2026-08-09'."""
    return (_dt.date.fromisoformat(iso) + _dt.timedelta(days=days)).isoformat()


def daterange(start: str, end: str, include_weekends: bool) -> list[str]:
    d0 = _dt.datetime.strptime(start, "%Y-%m-%d").date()
    d1 = _dt.datetime.strptime(end, "%Y-%m-%d").date()
    if d1 < d0:
        raise ValueError("End date is before start date.")
    out, cur = [], d0
    while cur <= d1:
        if include_weekends or cur.weekday() < 5:  # 5,6 = Sat,Sun
            out.append(cur.isoformat())
        cur += _dt.timedelta(days=1)
    return out


# --------------------------------------------------------------------------- #
#  Attendance parsing  (hours per date, from the portal's export file)
# --------------------------------------------------------------------------- #
def parse_total_minutes(value) -> Optional[int]:
    """
    Turn a 'Total Hours' cell into total MINUTES (no rounding).
      '10h 19m' -> 619     '9h 58m' -> 598     '7h 2m' -> 422
      '10:19'   -> 619     '10.5'   -> 630      'Weekend'/'--' -> None
    """
    if value is None:
        return None
    s = str(value).strip().lower()
    if not s or s in ("--", "-") or any(w in s for w in ("weekend", "leave", "absent", "holiday")):
        return None

    mh = re.search(r"(\d+)\s*h", s)
    mm = re.search(r"(\d+)\s*m", s)
    if mh or mm:
        h = int(mh.group(1)) if mh else 0
        m = int(mm.group(1)) if mm else 0
        return h * 60 + m

    if ":" in s:  # 10:19
        parts = s.split(":")
        try:
            return int(parts[0]) * 60 + int(parts[1])
        except (ValueError, IndexError):
            pass
    try:  # plain decimal, e.g. 10.5 hours
        return round(float(s) * 60)
    except ValueError:
        return None


def format_minutes(total: int) -> str:
    """Format minutes as a Jira duration string, e.g. 619 -> '10h 19m'."""
    total = max(0, int(total))
    h, m = divmod(total, 60)
    if h and m:
        return f"{h}h {m}m"
    if h:
        return f"{h}h"
    return f"{m}m"


def floor_to_hour(total_minutes: int) -> int:
    """
    Floor to the whole hour, but credit a half hour when the leftover minutes
    are 50 or more.
      9h13m -> 9h        9h45m -> 9h        9h50m -> 9h 30m
      9h58m -> 9h 30m    10h19m -> 10h      7h2m -> 7h
    """
    total_minutes = int(total_minutes)
    h, m = divmod(total_minutes, 60)
    return h * 60 + (30 if m >= 50 else 0)


def parse_any_date(value) -> Optional[str]:
    """Parse common date spellings into ISO 'YYYY-MM-DD'."""
    s = str(value).strip()
    if not s:
        return None
    fmts = ["%a, %d %b %Y", "%A, %d %b %Y", "%d %b %Y", "%d %B %Y",
            "%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y", "%d-%b-%Y", "%d-%m-%Y"]
    for f in fmts:
        try:
            return _dt.datetime.strptime(s, f).date().isoformat()
        except ValueError:
            continue
    return None


def _read_rows(path: str) -> list[list]:
    """Read a CSV or XLSX file into a list of rows (each a list of cells)."""
    ext = path.lower().rsplit(".", 1)[-1]
    if ext in ("xlsx", "xls"):
        try:
            import openpyxl  # noqa: WPS433
        except ImportError as exc:  # pragma: no cover
            raise RuntimeError(
                "Reading .xlsx needs openpyxl. Run: pip install openpyxl "
                "(or export the attendance as CSV instead)."
            ) from exc
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        return [[c.value for c in row] for row in ws.iter_rows()]
    with open(path, newline="", encoding="utf-8-sig") as fh:
        return [row for row in csv.reader(fh)]


LEAVE_KEYWORDS = (
    ("public holiday", "Public Holiday"),
    ("holiday", "Public Holiday"),
    ("annual leave", "Annual Leave"),
    ("casual leave", "Casual Leave"),
    ("compensatory", "Compensatory Leave"),
    ("short leave", "Short Leave"),
    ("leave", "Annual Leave"),   # generic fallback
)


def classify_row(row_text: str, minutes: Optional[int]) -> tuple:
    """
    Decide a day's category from its row text + parsed minutes.
    Returns (category, leave_label) where category is one of:
      'work' | 'leave' | 'weekend' | 'absent' | 'none'
    """
    t = row_text.lower()
    for needle, label in LEAVE_KEYWORDS:
        if needle in t:
            return "leave", label
    if minutes and minutes >= 1:
        return "work", None
    if "weekend" in t:
        return "weekend", None
    if "absent" in t:
        return "absent", None
    return "none", None


def load_attendance(path: str) -> dict[str, dict]:
    """Read an attendance export (CSV/XLSX) into the attendance dict."""
    return attendance_from_rows(_read_rows(path))


def attendance_from_rows(rows: list[list]) -> dict[str, dict]:
    """
    Turn attendance rows — from a file or scraped off the portal — into
        {iso_date: {"minutes": int|None, "category": str, "leave": str|None}}
    Auto-detects the Date and Total Hours columns; scans the whole row for
    leave / holiday / weekend labels.
    """
    if not rows:
        raise RuntimeError("Attendance data appears to be empty.")

    date_col = total_col = header_idx = None
    for idx, row in enumerate(rows[:10]):
        lowered = [str(c).strip().lower() if c is not None else "" for c in row]
        for j, cell in enumerate(lowered):
            if cell == "date" or cell.startswith("date"):
                date_col = j
            if "total" in cell and "hour" in cell:
                total_col = j
            elif cell == "work hours" and total_col is None:
                total_col = j
        if date_col is not None and total_col is not None:
            header_idx = idx
            break
    if date_col is None or total_col is None:
        raise RuntimeError(
            "Couldn't find 'Date' and 'Total Hours' columns in the export. "
            "Make sure you exported the attendance table with headers."
        )

    out: dict[str, dict] = {}
    for row in rows[header_idx + 1:]:
        if len(row) <= max(date_col, total_col):
            continue
        iso = parse_any_date(row[date_col])
        if not iso:
            continue
        minutes = parse_total_minutes(row[total_col])
        row_text = " ".join(str(c) for c in row if c is not None)
        category, leave = classify_row(row_text, minutes)
        if category == "none":
            continue
        out[iso] = {"minutes": minutes, "category": category, "leave": leave}
    return out


# --------------------------------------------------------------------------- #
#  Plan building  (one place, shared by the CLI and the web UI)
# --------------------------------------------------------------------------- #
STANDARD_DAY = 8 * 60      # minutes that make a full logged day
FULL_DAY = 9 * 60          # attendance at/above this counts as a long day


def duration_minutes(text) -> int:
    """'8h 30m' -> 510. A bare number is taken as minutes."""
    if isinstance(text, (int, float)):
        return max(0, int(text))
    s = str(text or "").strip().lower()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
    h = re.search(r"(\d+(?:\.\d+)?)\s*h", s)
    m = re.search(r"(\d+(?:\.\d+)?)\s*m", s)
    return int(round((float(h.group(1)) * 60 if h else 0) +
                     (float(m.group(1)) if m else 0)))


def classify_day(date: str, mode: str, attendance: dict,
                 leave_hours: str = "8h") -> dict:
    """
    What one day is worth, by the rules in section 4 of the README.

    Returns {"work":   minutes to log against the chosen ticket(s),
             "leaves": [(title, minutes), ...] bound for Planned Leaves}

    This is the single place those rules live: both plan builders reach it
    through remaining_day, so what a day is worth can never drift between the
    one-ticket-for-the-range path and the ticket-per-day one.
    """
    nothing = {"work": 0, "leaves": []}
    if mode == "flat":
        return {"work": STANDARD_DAY, "leaves": []}

    rec = attendance.get(date)
    if rec is None:
        return nothing                                 # day not in the sheet
    if rec["category"] == "absent":
        # Nothing automatic for an absent day: no hours to go by, and it must
        # never be filed as leave. What it is worth is what you type.
        return nothing
    if rec["category"] == "leave":
        return {"work": 0, "leaves": [(rec["leave"] or "Planned Leave",
                                       duration_minutes(leave_hours))]}
    if rec["category"] != "work" or not rec["minutes"]:
        return nothing

    t = floor_to_hour(rec["minutes"])
    if t < 1:
        return nothing
    if _dt.date.fromisoformat(date).weekday() >= 5:
        return {"work": t, "leaves": []}                # weekend: log the lot
    if t >= FULL_DAY:
        # Long day: a flat 8h in static mode, total - 1h otherwise.
        return {"work": STANDARD_DAY if mode == "static" else t - 60, "leaves": []}
    short = STANDARD_DAY - t
    return {"work": t, "leaves": [("Short Leave", short)] if short >= 30 else []}


def _logged_split(rec: Optional[dict], leave_parent: Optional[str],
                  target_work: int, target_leaves: list) -> tuple:
    """
    How much of a day's already-logged time counts against the work ticket,
    and how much against Planned Leaves.

    A day that is only work, or only leave, needs no guessing: all of it goes
    to the one bucket. When it is both -- a short day, worked plus a top-up --
    the sub-tasks sitting under the Planned Leaves issue say which is which.
    """
    total = int((rec or {}).get("minutes") or 0)
    if total <= 0:
        return 0, 0
    if not target_leaves:
        return total, 0
    if not target_work:
        return 0, total
    leave = 0
    if leave_parent:
        parents = (rec or {}).get("parents") or {}
        issues = (rec or {}).get("issues") or {}
        leave = int(parents.get(leave_parent) or 0) + int(issues.get(leave_parent) or 0)
    leave = min(leave, total)
    return total - leave, leave


def remaining_day(date: str, mode: str, attendance: dict, *,
                  logged: Optional[dict] = None,
                  leave_parent: Optional[str] = None,
                  leave_hours: str = "8h") -> dict:
    """
    What still needs logging on one date: what the day is worth (classify_day)
    less whatever your Jira worklogs already cover for it.

    Returns classify_day's shape, plus what it was worked out from:
        {"work", "leaves",                       # what is left to create
         "target_work", "target_leaves",         # what the day is worth
         "logged_work", "logged_leave", "logged"}

    Without a `logged` map this is classify_day exactly, so a caller that does
    not care about existing worklogs behaves as it always did. A day that is
    already covered comes back with nothing left to do, which is what keeps the
    same hours from being logged twice.
    """
    day = classify_day(date, mode, attendance, leave_hours)
    target_work, target_leaves = day["work"], day["leaves"]
    rec = (logged or {}).get(date) or {}
    done_work, done_leave = _logged_split(rec, leave_parent, target_work,
                                          target_leaves)
    leaves, spare = [], done_leave
    for title, mins in target_leaves:
        take = min(mins, spare)
        spare -= take
        if mins - take >= 1:
            leaves.append((title, mins - take))
    return {"work": max(0, target_work - done_work), "leaves": leaves,
            "target_work": target_work, "target_leaves": target_leaves,
            "logged_work": done_work, "logged_leave": done_leave,
            "logged": int(rec.get("minutes") or 0)}


def split_minutes(total: int, shares: list) -> list[int]:
    """
    Divide a day's minutes by percentage without losing or inventing a minute.

    Whole minutes each, largest remainder first, so 570 split 60/40 comes back
    as 342 + 228 and not 341 + 227. Shares that don't add to 100 are treated as
    weights; all-zero (or missing) shares split the day evenly.
    """
    total = max(0, int(total))
    if not shares:
        return []
    weights = [max(0.0, float(s or 0)) for s in shares]
    pool = sum(weights)
    if pool <= 0:
        weights, pool = [1.0] * len(shares), float(len(shares))
    exact = [total * w / pool for w in weights]
    out = [int(x) for x in exact]
    order = sorted(range(len(out)), key=lambda i: exact[i] - out[i], reverse=True)
    for i in range(total - sum(out)):
        out[order[i % len(order)]] += 1
    return out


def _work_entry(date: str, parent: str, summary_base: str, task_value: str,
                minutes: int, share=None, label=None) -> dict:
    entry = {"date": date, "kind": "work", "parent": parent,
             "title": summary_base, "task_value": task_value,
             "estimate": format_minutes(minutes), "logged": format_minutes(minutes)}
    if share is not None:
        entry["share"] = share
    # The title this day was given when it was assigned. Days grouped into one
    # sub-task can have been given different ones, so each keeps its own and
    # the sub-task offers them as the names it could go by.
    if label is not None:
        entry["label"] = label
    return entry


def _leave_entry(date: str, leave_parent: Optional[str], title: str,
                 minutes: int) -> dict:
    return {"date": date, "kind": "leave", "parent": leave_parent,
            "title": title, "task_value": title,
            "estimate": format_minutes(minutes), "logged": format_minutes(minutes)}


def build_plan(*, start: str, end: str, mode: str, attendance: dict,
               parent: str, leave_parent: Optional[str], summary_base: str,
               task_value: str, leave_hours: str = "8h",
               include_weekends: bool = False,
               logged: Optional[dict] = None) -> list[dict]:
    """
    Turn a date range + the attendance sheet into the sub-tasks to create,
    with one ticket carrying the whole range.

    Modes:
      'attendance'  long weekday (>= 9h) logs the floored total minus 1h
      'static'      long weekday (>= 9h) logs a flat 8h instead
      'flat'        no sheet at all: a plain 8h on every weekday

    Everything else is identical in 'attendance' and 'static':
      short weekday (< 9h) -> log what was worked, plus a Short Leave under
      Planned Leaves for the rest of the 8h; weekend with hours -> log the
      full total; leave / public holiday -> leave_hours under Planned Leaves.

    `logged` is what Jira already holds for the range (JiraClient.my_worklogs).
    Pass it and every day is trimmed by what it already has -- a day logged in
    full drops out, a day logged short keeps only the difference.
    """
    entries: list[dict] = []
    for d in daterange(start, end, include_weekends or mode != "flat"):
        day = remaining_day(d, mode, attendance, logged=logged,
                            leave_parent=leave_parent, leave_hours=leave_hours)
        if day["work"]:
            entries.append(_work_entry(d, parent, summary_base, task_value,
                                       day["work"]))
        for title, minutes in day["leaves"]:
            entries.append(_leave_entry(d, leave_parent, title, minutes))
    return entries


def _row_minutes(rows: list, room: int) -> list[int]:
    """
    How many minutes each row of a day gets.

    Rows that state their own `hours` are taken at their word, which is what
    lets a day be assigned in parts - 5h to one ticket now, the other 3h to
    another later - rather than always being divided up in one go. Anything
    over what the day still owes is scaled back down so a day can never
    overrun itself.

    Rows with no hours fall back to `share`, exactly as they always did, so a
    day assigned in one pass (and every CLI caller) behaves as before.
    """
    stated = [duration_minutes(r.get("hours")) if r.get("hours") else None
              for r in rows]
    if not any(m is not None for m in stated):
        return split_minutes(room, [r.get("share") for r in rows])
    # Mixed rows: the ones that named their hours keep them, and whatever is
    # left of the day is shared out among the rest by their shares.
    fixed = sum(m for m in stated if m)
    if fixed >= room:
        # They ask for at least the whole day: give them the day, in proportion.
        return split_minutes(room, [m or 0 for m in stated])
    rest = [i for i, m in enumerate(stated) if m is None]
    out = [m or 0 for m in stated]
    if rest:
        share = split_minutes(room - fixed, [rows[i].get("share") for i in rest])
        for i, mins in zip(rest, share):
            out[i] = mins
    return out


def _absent_entries(date: str, rows: list, day: dict, already: list) -> list:
    """
    What to create for a day the sheet marks absent.

    Nothing at all unless a row claims it. When one does, the hours are the
    ones typed against it, less anything Jira already holds for that date, so
    the same promise applies here as everywhere else: hours cannot go in twice.
    """
    if not rows:
        return []
    wanted = [duration_minutes(r.get("hours")) for r in rows]
    total = sum(wanted)
    if total <= 0:
        return []
    room = max(0, total - int(day.get("logged_work") or 0))
    if not room:
        already.append(date)
        return []
    out = []
    for row, mins in zip(rows, split_minutes(room, wanted)):
        if mins <= 0:
            continue
        task = (row.get("task_value") or "").strip()
        title = (row.get("title") or "").strip()
        out.append(_work_entry(date, (row.get("parent") or "").strip(),
                               f"{task}-{title}" if title else task, task, mins,
                               share=row.get("share"), label=title))
    return out


def build_plan_items(*, start: str, end: str, mode: str, attendance: dict,
                     items: dict, leave_parent: Optional[str],
                     leave_hours: str = "8h",
                     logged: Optional[dict] = None) -> tuple:
    """
    The same hour rules as build_plan, but each day's hours are split across
    whichever ticket(s) were chosen for that day.

    `items` maps an ISO date to the rows chosen for it:
        {"2026-07-28": [{"parent", "task_value", "title", "share"}, ...]}
    where `share` is that row's percentage of the day. Leave and public-holiday
    days need no row: they go to Planned Leaves exactly as they always did.

    A day the sheet marks **absent** has no hours of its own, so it is only
    ever created when a row claims it, and then for the hours that row carries
    (`hours`) rather than a share of a day. Whatever Jira already holds for
    that date still comes off the top.

    Returns (plan, unassigned, already). `unassigned` lists the working days in
    range that no row claimed, so the caller can say so rather than quietly
    skip them; `already` lists the days Jira has covered in full, which are not
    the user's to assign at all.
    """
    plan: list[dict] = []
    unassigned: list[str] = []
    already: list[str] = []
    for d in daterange(start, end, mode != "flat"):
        day = remaining_day(d, mode, attendance, logged=logged,
                            leave_parent=leave_parent, leave_hours=leave_hours)
        rows = items.get(d) or []
        if (attendance.get(d) or {}).get("category") == "absent":
            plan.extend(_absent_entries(d, rows, day, already))
            continue
        if day["target_work"] and not day["work"]:
            already.append(d)               # every hour of it is already in Jira
        if day["work"]:
            if not rows:
                unassigned.append(d)
            else:
                minutes = _row_minutes(rows, day["work"])
                for row, mins in zip(rows, minutes):
                    if mins <= 0:
                        continue                    # a 0% row creates nothing
                    task = (row.get("task_value") or "").strip()
                    title = (row.get("title") or "").strip()
                    plan.append(_work_entry(
                        d, (row.get("parent") or "").strip(),
                        f"{task}-{title}" if title else task, task, mins,
                        share=row.get("share"), label=title))
        for title, mins in day["leaves"]:
            plan.append(_leave_entry(d, leave_parent, title, mins))
    return plan, unassigned, already


# --------------------------------------------------------------------------- #
#  Grouping and packing  (many days -> one sub-task, days kept whole)
# --------------------------------------------------------------------------- #
PACK_CAP = 24 * 60         # the most one sub-task may hold, in minutes


def title_label(title: str, task_value: str) -> str:
    """
    The part of a summary base the user typed, with the Task taken off.
        ('Development-Edit Invoice Screen', 'Development') -> 'Edit Invoice Screen'
        ('Development', 'Development')                     -> ''
    """
    title = title or ""
    if not task_value:
        return title
    if title == task_value:
        return ""
    prefix = f"{task_value}-"
    return title[len(prefix):] if title.startswith(prefix) else title


def summary_base(task_value: str, label) -> str:
    """The inverse: 'Development' + 'Edit Invoice Screen' -> both, joined."""
    label = str(label or "").strip()
    if not task_value:
        return label
    return f"{task_value}-{label}" if label else task_value


def chunk_summary(title: str, dates: list = None) -> str:
    """
    The sub-task title: its Task, plus your title if you gave one.

        'Development'                          (no title)
        'Development-Invoice totals rework'    (titled)

    No dates. A sub-task carries a work log per day it holds, and those say
    which days far better than a name ever did. `dates` is accepted and
    ignored so older callers keep working.
    """
    return title


def title_options(entries: list) -> list:
    """
    The titles the days of one sub-task were given, the biggest share first.

        [{"label": "Invoice totals", "days": 3, "minutes": 1440, "hours": "24h"}]

    A sub-task holding days titled differently can go by any of them, so the
    caller can offer the list and let the user say which - or none.
    """
    seen: dict = {}
    for e in entries:
        label = str(e.get("label") or "").strip()
        if not label:
            continue
        row = seen.setdefault(label, {"label": label, "days": 0, "minutes": 0})
        row["days"] += 1
        row["minutes"] += duration_minutes(e.get("logged"))
    out = sorted(seen.values(), key=lambda r: (-r["minutes"], r["label"]))
    for row in out:
        row["hours"] = format_minutes(row["minutes"])
    return out


def _chunk(entries: list, *, kind: str, parent, task_value: str,
           title: str = None) -> dict:
    """One sub-task built from whole day entries, sized to their real total."""
    logs = [{"date": e["date"], "hours": e["logged"],
             "minutes": duration_minutes(e["logged"])}
            for e in entries if e.get("logged")]
    minutes = sum(l["minutes"] for l in logs)
    dates = [e["date"] for e in entries]
    titles = title_options(entries)
    if title is None:
        # No name imposed: go by the title that covers most of the sub-task,
        # and let the caller pick another from `titles`.
        label = titles[0]["label"] if titles else ""
        title = summary_base(task_value, label)
    else:
        label = title_label(title, task_value)
    return {"kind": kind, "parent": parent, "title": title,
            "task_value": task_value,
            # The title on its own, so a sub-task can be renamed later without
            # unpicking the summary string, and the titles it may go by.
            "label": label, "titles": titles,
            "date": dates[0], "dates": dates, "days": len(dates),
            "estimate": format_minutes(minutes), "logged": format_minutes(minutes),
            "logs": logs, "summary": chunk_summary(title)}


def pack_plan(plan: list, cap_minutes: int = None) -> list:
    """
    Turn a day-by-day plan into the sub-tasks that will actually be created.

Entries for the same ticket and Task collect together and their days are
    packed, in date order, into sub-tasks of at most `cap_minutes`.

    Whatever titles those days were given come along: each sub-task lists them
    in `titles` and takes the biggest as its own `label`, so the caller can
    offer the choice rather than guessing.
    A day is never split across two sub-tasks: when the next whole day would
    take a sub-task over the cap, that sub-task closes at its real total and
    the next one starts. A single day worth more than the cap is a sub-task on
    its own - hence a *soft* cap.

    Leave and public holidays are never packed: grouping them would lose which
    day was which kind of leave, so each stays its own entry.

    Each returned entry carries `logs`, one per day, so the run puts the hours
    back on the days they were worked.
    """
    # Read the cap now, not when this was defined, so editing PACK_CAP - or
    # passing --pack-cap - actually takes effect.
    cap_minutes = PACK_CAP if cap_minutes is None else cap_minutes
    buckets: dict = {}
    order: list = []
    leaves: list = []
    for e in plan:
        if e.get("kind") != "work":
            leaves.append(_chunk([e], kind=e.get("kind", "leave"),
                                 parent=e.get("parent"), title=e["title"],
                                 task_value=e.get("task_value", "")))
            continue
        key = (e.get("parent") or "", e.get("task_value") or "")
        if key not in buckets:
            buckets[key] = []
            order.append(key)
        buckets[key].append(e)

    out: list = []
    for key in order:
        parent, task_value = key
        entries = sorted(buckets[key], key=lambda x: x["date"])
        chunk: list = []
        total = 0
        for e in entries:
            mins = duration_minutes(e.get("logged"))
            if chunk and total + mins > cap_minutes:
                out.append(_chunk(chunk, kind="work", parent=parent,
                                  task_value=task_value))
                chunk, total = [], 0
            chunk.append(e)
            total += mins
        if chunk:
            out.append(_chunk(chunk, kind="work", parent=parent,
                              task_value=task_value))

    # Work first, grouped ticket by ticket the way the preview reads it, then
    # the leave days in date order.
    out.sort(key=lambda c: (c["parent"] or "", c["task_value"], c["date"]))
    leaves.sort(key=lambda c: (c["date"], c["title"]))
    return out + leaves


def plan_logs(entry: dict) -> list:
    """
    The day-by-day work logs of one plan entry.

    Older entries (a retry from a page that predates grouping, or the CLI's
    one-day-per-sub-task shape) carry no `logs`, so their single day is
    presented in the same form and every caller can loop.
    """
    logs = entry.get("logs")
    if logs:
        return [l for l in logs if l.get("hours")]
    if not entry.get("logged"):
        return []
    return [{"date": entry["date"], "hours": entry["logged"],
             "minutes": duration_minutes(entry["logged"])}]


def entry_base(entry: dict) -> str:
    """
    An entry's summary base: its Task plus whatever title it has been given.

    A caller that renamed a sub-task sends the new `label` back, and this is
    where that takes effect - so the name Jira gets is worked out here, not
    trusted from whatever string came in with it.
    """
    if "label" in entry:
        return (summary_base(entry.get("task_value") or "", entry.get("label"))
                or entry.get("title") or "")
    return entry.get("title") or ""


def entry_summary(entry: dict) -> str:
    """The sub-task title to create for an entry, grouped or not."""
    if "label" in entry:
        return chunk_summary(entry_base(entry))
    return entry.get("summary") or entry.get("title") or ""


# --------------------------------------------------------------------------- #
#  Interactive prompts
# --------------------------------------------------------------------------- #
def clean_dropped_path(raw: str) -> str:
    """
    Normalise a path that was typed OR drag-and-dropped onto the terminal.
    Drag-and-drop usually pastes the path wrapped in quotes (and PowerShell
    may prefix it with '& '); some shells escape spaces with backslashes.
    """
    s = raw.strip()
    if s.startswith("& "):            # PowerShell "& 'C:\\path'"
        s = s[2:].strip()
    # Strip one layer of matching surrounding quotes.
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
        s = s[1:-1]
    s = s.strip().strip("'\"")
    s = s.replace("\\ ", " ")         # unescape spaces (macOS/Linux drag)
    return os.path.expanduser(s)


def browse_for_file() -> Optional[str]:
    """Open a native file-picker dialog. Returns a path or None if unavailable."""
    try:
        import tkinter as tk
        from tkinter import filedialog
    except Exception:  # noqa: BLE001
        return None
    try:
        root = tk.Tk()
        root.withdraw()
        root.attributes("-topmost", True)
        path = filedialog.askopenfilename(
            title="Select the attendance export",
            filetypes=[("Excel/CSV", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        root.destroy()
        return path or None
    except Exception:  # noqa: BLE001
        return None


def ask_attendance_file(current: Optional[str] = None) -> str:
    """
    Get the attendance file. Accepts a drag-and-dropped file (paste the path),
    a typed path, or — on pressing Enter — opens a file-browse dialog.
    Re-prompts until a real .xlsx/.xls/.csv file is given.
    """
    if current:
        return current
    while True:
        raw = input(
            "\nDrag & drop the attendance sheet here and press Enter "
            "(or press Enter to browse): "
        )
        path = clean_dropped_path(raw)
        if not path:                       # empty -> open a browse dialog
            picked = browse_for_file()
            if picked:
                path = picked
            else:
                print("  No file chosen. Drag the .xlsx onto the window or type its path.")
                continue
        if not os.path.isfile(path):
            print(f"  Can't find a file at: {path}")
            continue
        if not path.lower().endswith((".xlsx", ".xls", ".csv")):
            print("  Please provide a .xlsx (or .csv) attendance export.")
            continue
        return path


def prompt(label: str, current: Optional[str] = None, required: bool = True) -> str:
    if current:
        return current
    while True:
        val = input(f"{label}: ").strip()
        if val or not required:
            return val
        print("  (required)")


def choose_from_menu(title: str, options: list[str], default: Optional[str] = None) -> str:
    """Print a numbered menu and return the option the user selects.

    If `default` matches one of the options, the user can just press Enter
    to accept it.
    """
    print(f"\n{title}")
    default_idx = None
    for i, opt in enumerate(options, 1):
        marker = ""
        if default and opt.strip().lower() == default.strip().lower():
            default_idx = i
            marker = "  (default)"
        print(f"  {i:>2}) {opt}{marker}")
    hint = f" (Enter for {options[default_idx - 1]})" if default_idx else ""
    while True:
        raw = input(f"Select a number [1-{len(options)}]{hint}: ").strip()
        if not raw and default_idx:
            return options[default_idx - 1]
        if raw.isdigit() and 1 <= int(raw) <= len(options):
            return options[int(raw) - 1]
        print("  invalid choice, try again")


# --------------------------------------------------------------------------- #
#  Main
# --------------------------------------------------------------------------- #
def main(argv: Optional[list[str]] = None) -> int:
    p = argparse.ArgumentParser(description="Create daily 8h sub-tasks, log them, mark Done.")
    p.add_argument("--base-url", default="https://tracking.i2cinc.com/",
                   help="Jira base URL (default: https://tracking.i2cinc.com/)")
    p.add_argument("--username", help="Username (or email on Cloud). Blank = treat password as PAT.")
    p.add_argument("--parent", help="Parent issue key (skip sprint selection), e.g. CR-10835")
    p.add_argument("--sprint", help="Sprint name, e.g. ST12-26.8 (pick a parent issue from it)")
    p.add_argument("--board", type=int, default=None,
                   help="Agile board id for the sprint lookup. Optional — the "
                        "board is detected from the sprint name by default.")
    p.add_argument("--start", help="Start date YYYY-MM-DD")
    p.add_argument("--end", help="End date YYYY-MM-DD")
    p.add_argument("--summary", help='Base sub-task title, e.g. "Development"')
    p.add_argument("--status-path", default="In Progress,Done",
                   help="Comma-separated status path to walk (default: 'In Progress,Done')")
    p.add_argument("--task-value", default=None,
                   help="Value for the 'Task' dropdown. Defaults to the sub-task title you enter.")
    p.add_argument("--original-estimate", default="9h",
                   help="Original estimate when NOT using an attendance file (default: 9h)")
    p.add_argument("--hours", default="8h",
                   help="Time logged when NOT using an attendance file (default: 8h)")
    p.add_argument("--attendance-file",
                   help="Attendance export (CSV/XLSX). Estimate = total hours; logged = "
                        "total-1h on weekdays, full total on weekends.")
    p.add_argument("--static-hours", action="store_true",
                   help="Log a flat 8h on full days. With --attendance-file the "
                        "sheet still drives short days, weekends and leave; "
                        "without one, every weekday simply gets 8h.")
    p.add_argument("--leave-parent", default=None,
                   help="Parent issue for leave/holiday days. Default: auto-detect the "
                        "sprint issue whose summary contains 'Planned Leaves'.")
    p.add_argument("--leave-hours", default="8h",
                   help="Estimate and logged time for leave/holiday sub-tasks (default: 8h)")
    p.add_argument("--no-pack", action="store_true",
                   help="One sub-task per day, the way it worked before. By "
                        "default the days for a ticket+Task are grouped into "
                        "sub-tasks of up to --pack-cap, each carrying one work "
                        "log per day.")
    p.add_argument("--pack-cap", type=float, default=24.0,
                   help="Most hours one grouped sub-task may hold (default: 24). "
                        "A day is never split, so a longer single day stands alone.")
    p.add_argument("--ignore-logged", action="store_true",
                   help="Log the full day even where Jira already has hours "
                        "for it. By default your own worklogs are read back "
                        "and each day is trimmed by what it already has.")
    p.add_argument("--no-assign-self", action="store_true",
                   help="Do NOT assign the created sub-tasks to yourself")
    p.add_argument("--no-check-all", action="store_true",
                   help="Do NOT auto-tick the Development Checklist boxes")
    p.add_argument("--include-weekends", action="store_true", help="Also create tasks on Sat/Sun")
    p.add_argument("--no-verify-ssl", action="store_true", help="Disable SSL verification")
    p.add_argument("--dry-run", action="store_true", help="Preview only; make no changes")
    args = p.parse_args(argv)

    print("=== Jira timesheet automation ===\n")

    base_url = prompt("Jira base URL", args.base_url)

    # --- connect + verify login (retry a few times) ---
    client = None
    for attempt in range(1, 4):
        username = (args.username if args.username is not None
                    else input("Username (blank if using a PAT): ").strip())
        try:
            password = getpass.getpass("Password / API token / PAT: ")
        except (EOFError, KeyboardInterrupt):
            print("\nCancelled.")
            return 1
        if not password:
            print("  Password can't be empty.\n")
            args.username = None            # let them retype the username too
            continue
        try:
            client = JiraClient(base_url, username, password,
                                verify_ssl=not args.no_verify_ssl)
            who = client.verify_login()
            print(f"\nLogged in as {who}\n")
            break
        except LoginError as exc:
            print(f"\n{exc}\n", file=sys.stderr)
            client = None
            args.username = None            # re-prompt username on next try
            if attempt < 3:
                print(f"Let's try again ({attempt}/3).\n")
        except Exception as exc:  # noqa: BLE001
            print(f"\nUnexpected error while logging in: {exc}\n", file=sys.stderr)
            client = None
            args.username = None
    if client is None:
        print("Could not log in after 3 attempts. Exiting.", file=sys.stderr)
        return 2

    # Loop so the user can log several tasks/issues without re-entering the password.
    while True:
        try:
            rc = run_session(client, args)
        except KeyboardInterrupt:
            print("\nCancelled this session.")
            return 1
        if args.dry_run or rc == 2:
            return rc
        try:
            again = input("\nLog another task? [y/n]: ").strip().lower()
        except (EOFError, KeyboardInterrupt):
            print("\nDone. Bye.")
            return rc
        if again != "y":
            print("Done. Bye.")
            return rc


def run_session(client: "JiraClient", args) -> int:
    """One create-and-log session. Returns 0 ok, 1 partial, 2 error/abort."""
    # --- pick the parent issue (from a sprint, unless --parent given) ---
    sprint_issues: list[tuple] = []
    if args.parent:
        parent = args.parent
    else:
        sprint = prompt("Sprint (e.g. ST12-26.8)", args.sprint)
        try:
            sprint_issues, sprint_meta = client.search_sprint(sprint, board_id=args.board)
        except Exception as exc:  # noqa: BLE001
            print(f"Error searching sprint: {exc}", file=sys.stderr)
            return 2
        if sprint_meta.get("board_id") and not args.board:
            print(f"  (sprint '{sprint_meta['sprint_name']}' found on board "
                  f"{sprint_meta['board_id']})")
        if not sprint_issues:
            print(f"No issues found in sprint '{sprint}'.", file=sys.stderr)
            return 2
        labels = [f"{k}  -  {s}" for k, s, *_ in sprint_issues]
        chosen = choose_from_menu("Which issue do you want to create sub-tasks under?", labels)
        parent = sprint_issues[labels.index(chosen)][0]

    # Resolve the leave/holiday parent: explicit flag, else auto-detect from
    # the sprint the issue whose summary mentions "Planned Leaves".
    leave_parent = args.leave_parent
    if not leave_parent:
        for k, s, *_ in sprint_issues:
            if "planned leave" in s.lower():
                leave_parent = k
                break

    start = prompt("Start date (YYYY-MM-DD)", args.start)
    end = prompt("End date (YYYY-MM-DD)", args.end)

    # --- resolve the selected issue's project + sub-task create context ---
    try:
        ctx = client.get_parent_context(parent)
        work_project = ctx["project_key"]
        work_subtask_id, create_fields = client.get_project_subtask_context(work_project)
    except Exception as exc:  # noqa: BLE001
        print(f"Error: {exc}", file=sys.stderr)
        return 2

    # --- choose the Task type (pulled live from Jira's dropdown) ---
    task_fid, task_opts = client.get_task_options(create_fields)
    task_values = [v for v, _ in task_opts]

    if args.task_value:
        task_value = args.task_value
    elif task_values:
        task_value = choose_from_menu("Which Task do you want to create sub-tasks for?",
                                      task_values)
    else:
        task_value = prompt('Task value (couldn\'t load list)', None) or "Development"

    # Ask for a descriptive title (optional). Final work title:
    #   with a title  -> "<Task>-<your title>-<date>"
    #   skipped       -> "<Task>-<date>"
    if args.summary is not None:
        user_title = args.summary.strip()
    else:
        print(f'\nEnter a title for the {task_value} task, or press Enter to skip.')
        print(f'  with title : {task_value}-<your title>-<date>')
        print(f'  if skipped : {task_value}-<date>')
        user_title = input("Title (optional): ").strip()
    summary_base = f"{task_value}-{user_title}" if user_title else task_value

    # --- how should hours be set? ---
    # 'attendance' and 'static' both read the sheet; they differ only in what a
    # long weekday logs (total - 1h vs a flat 8h). 'flat' skips the sheet.
    if args.static_hours and not args.attendance_file:
        mode = "flat"
    elif args.attendance_file:
        mode = "static" if args.static_hours else "attendance"
    else:
        pick = choose_from_menu(
            "How should the hours be set?",
            ["From attendance (TotalHours)",
             "From attendance, but a flat 8h on full days",
             "Static 8 hours per weekday (no sheet)"],
        )
        mode = ("attendance" if pick.startswith("From attendance (") else
                "static" if pick.startswith("From attendance,") else "flat")

    attendance: dict[str, dict] = {}
    if mode != "flat":
        att_path = ask_attendance_file(args.attendance_file)
        args.attendance_file = att_path      # remember it for the next round
        try:
            attendance = load_attendance(att_path)
        except Exception as exc:  # noqa: BLE001
            print(f"Error reading attendance file: {exc}", file=sys.stderr)
            return 2
        if not attendance:
            print("No usable rows found in the attendance file.", file=sys.stderr)
            return 2

    status_path = [s.strip() for s in args.status_path.split(",") if s.strip()]

    # --- what is already on your Jira timesheet for this range? ---
    # A day your own worklogs already cover is left alone; a day logged short
    # gets only the difference. --ignore-logged turns the check off.
    logged: dict = {}
    if not args.ignore_logged:
        try:
            logged = client.my_worklogs(start, end)
        except Exception as exc:  # noqa: BLE001 - a check, not the job itself
            print(f"  (couldn't read your existing worklogs: {exc})",
                  file=sys.stderr)
            print("   carrying on without them - nothing will be subtracted.")
        else:
            if logged:
                have = sum(int(r["minutes"]) for r in logged.values())
                print(f"  ({len(logged)} day(s) already logged in Jira, "
                      f"{format_minutes(have)} in total)")

    try:
        plan = build_plan(start=start, end=end, mode=mode, attendance=attendance,
                          parent=parent, leave_parent=leave_parent,
                          summary_base=summary_base, task_value=task_value,
                          leave_hours=args.leave_hours,
                          include_weekends=args.include_weekends,
                          logged=logged)
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        return 2
    # Group the days into the sub-tasks that will actually be created.
    day_count = len({e["date"] for e in plan})
    if not args.no_pack:
        plan = pack_plan(plan, int(round(max(0.5, args.pack_cap) * 60)))
    else:
        plan = [_chunk([e], kind=e["kind"], parent=e["parent"],
                       title=e["title"], task_value=e["task_value"])
                for e in plan]
    dates = daterange(start, end, args.include_weekends or mode != "flat")
    # The same view of each day the plan was built from, for the report below.
    day_view = {d: remaining_day(d, mode, attendance, logged=logged,
                                 leave_parent=leave_parent,
                                 leave_hours=args.leave_hours) for d in dates}
    covered = [d for d, v in day_view.items()
               if (v["target_work"] or v["target_leaves"])
               and not (v["work"] or v["leaves"])]
    topped_up = [d for d, v in day_view.items()
                 if v["logged"] and (v["work"] or v["leaves"])]
    need_leave_ctx = any(e["kind"] == "leave" for e in plan)

    # Resolve the leave parent's context only if we actually have leave days.
    leave_project = leave_subtask_id = leave_create_fields = None
    if need_leave_ctx:
        if not leave_parent:
            print("There are leave/holiday days, but no 'Planned Leaves' issue was "
                  "found in the sprint. Pass --leave-parent <ISSUE-KEY> to set it.",
                  file=sys.stderr)
            return 2
        try:
            lctx = client.get_parent_context(leave_parent)
            leave_project = lctx["project_key"]
            leave_subtask_id, leave_create_fields = \
                client.get_project_subtask_context(leave_project)
        except Exception as exc:  # noqa: BLE001
            print(f"Error resolving leave parent {leave_parent}: {exc}", file=sys.stderr)
            return 2

    # --- base fields (assignee) shared by everything ---
    base_assignee = {}
    if not args.no_assign_self and client.assignee_ref:
        base_assignee["assignee"] = client.assignee_ref
    assign_note = "assign to me" if base_assignee else "unassigned"

    def build_day_fields(entry: dict) -> dict:
        """Assemble the create fields for one plan entry (its own project's Task field)."""
        fields = dict(base_assignee)
        if entry["estimate"]:
            fields["timetracking"] = {"originalEstimate": entry["estimate"]}
        cf = create_fields if entry["kind"] == "work" else (leave_create_fields or {})
        if entry["task_value"]:
            fid, ref = client.resolve_task_field(cf, entry["task_value"])
            if fid and ref:
                fields[fid] = ref
        return fields

    # --- plan summary ---
    n_work = sum(1 for e in plan if e["kind"] == "work")
    n_leave = sum(1 for e in plan if e["kind"] == "leave")
    check_note = "checklist auto-ticked" if not args.no_check_all else "checklist left blank"
    print(f"\nPlan: {len(plan)} sub-tasks from {day_count} day(s)  "
          f"({n_work} work under {parent}, "
          f"{n_leave} leave/holiday under {leave_parent or '(none)'}).")
    if not args.no_pack:
        print(f"      grouped by ticket + Task, up to "
              f"{format_minutes(int(round(args.pack_cap * 60)))} a sub-task, "
              f"days kept whole")
    print(f"      create: {assign_note} | Task = {task_value} | {check_note}")
    if covered:
        print(f"      already logged in Jira, left alone: {len(covered)} day(s)")
    if topped_up:
        print(f"      logged short in Jira, topping up: {len(topped_up)} day(s)")
    print(f"      status path: {' -> '.join(status_path)}\n")

    if args.dry_run:
        for e in plan:
            tag = "LEAVE" if e["kind"] == "leave" else "work "
            print(f"  [{tag}] '{entry_summary(e)}' under {e['parent']} "
                  f"| est {e['estimate']} | log {e['logged'] or '—'}"
                  + (f" | Task={e['task_value']}" if e['task_value'] else ""))
            if e.get("days", 1) > 1:
                inside = "  ".join(f"{l['date'][5:]} {l['hours']}"
                                   for l in plan_logs(e))
                print(f"          {e['days']} work logs: {inside}")
        for d in topped_up:
            v = day_view[d]
            worth = v["target_work"] + sum(m for _, m in v["target_leaves"])
            print(f"  [top up] {d}  {format_minutes(v['logged'])} of "
                  f"{format_minutes(worth)} already in Jira - logging the rest")
        if covered:
            print(f"  already logged, leaving {len(covered)} day(s) alone: "
                  f"{', '.join(covered)}")
        # A sub-task now covers several days, so what it holds is in its work
        # logs - reading only `date` would report its own days as skipped.
        planned = {l["date"] for e in plan for l in plan_logs(e)}
        skipped = [d for d in dates if d not in planned and d not in covered]
        if skipped:
            print(f"  skipping {len(skipped)} day(s): {', '.join(skipped)}")
        print("\n(dry run - nothing was changed)")
        return 0

    confirm = input(f"Proceed with {len(plan)} sub-tasks? [y/n]: ").strip().lower()
    if confirm != "y":
        print("Aborted.")
        return 0

    # --- run ---
    results: list[DayResult] = []
    for e in plan:
        logs = plan_logs(e)
        res = DayResult(date=f"{e['date']} ({e['kind']})",
                        hours=e["logged"], days=len(logs))
        try:
            project = work_project if e["kind"] == "work" else leave_project
            subtask_id = work_subtask_id if e["kind"] == "work" else leave_subtask_id
            key = client.create_subtask(project, e["parent"], subtask_id,
                                        entry_summary(e),
                                        extra_fields=build_day_fields(e))
            res.created_key = key
            for log in logs:
                client.log_work(key, log["date"], log["hours"], comment=e["title"])
            res.logged = bool(logs)
            client.move_through(key, status_path,
                                task_value=e["task_value"] or task_value,
                                check_all=not args.no_check_all,
                                project=project)
            res.done = True
        except Exception as exc:  # noqa: BLE001
            res.error = str(exc)
        results.append(res)
        print("  " + str(res))

    ok = sum(r.ok for r in results)
    print(f"\nDone: {ok}/{len(results)} sub-tasks fully created, logged, and closed "
          f"({sum(r.days for r in results if r.logged)} work log(s)).")
    return 0 if ok == len(results) else 1


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except KeyboardInterrupt:
        print("\nCancelled.")
        raise SystemExit(1)
