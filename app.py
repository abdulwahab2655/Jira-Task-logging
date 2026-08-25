#!/usr/bin/env python3
"""
app.py — Local web frontend for the Jira timesheet utility.

Run it, then open the browser it points you to. Everything stays on your
machine: your login, your Jira, your attendance file. It reuses the exact
logic from jira_logging_utility.py, so behaviour is identical to the CLI.

    pip install flask requests openpyxl selenium
    python app.py
    -> open http://127.0.0.1:5000

"Sign in with Microsoft" is fully automatic: the backend drives the browser
sign-in and picks the Jira session up by itself (see sso_login.py). Selenium is
what makes that work; without it, only username/password/PAT sign-in is available.
"""

from __future__ import annotations

import os
import tempfile
from concurrent.futures import ThreadPoolExecutor
import threading
import webbrowser

from flask import Flask, jsonify, request, send_from_directory

import attendance_portal
import jira_credentials
import jira_logging_utility as core
import sso_login

# Serve files from the folder this script lives in (works whether index.html
# sits next to app.py or inside a 'static' subfolder).
HERE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)

# How much of the run happens at once. Jira answers in about 0.3s, and a
# sub-task is a create, a work log per day and two transitions - almost all of
# it waiting, so it is worth overlapping. Sub-tasks are independent of each
# other, and so are the days inside one: that is the second dimension.
RUN_WORKERS = 8            # sub-tasks on the go at once
LOG_WORKERS = 4            # days of one sub-task logged at once


def _find_index() -> tuple:
    """Return (directory, filename) for index.html, checking ./ and ./static."""
    for folder in (HERE, os.path.join(HERE, "static")):
        candidate = os.path.join(folder, "index.html")
        if os.path.isfile(candidate):
            return folder, "index.html"
    return HERE, "index.html"

# Single-user local app: keep one logged-in client + a little state in memory.
STATE: dict = {
    "client": None,
    "sprint_issues": [],     # [(key, summary)]
    "leave_parent": None,
    "attendance": {},        # {iso: {minutes, category, leave}}
    "logged": {},            # what Jira already has: {iso: {minutes, parents, issues}}
    "logged_range": "",      # the 'start|end' those worklogs were read for
    "projects": {},          # cache: parent_key -> project_key (one cheap call)
    "contexts": {},          # cache: project_key -> (subtask_id, create_fields)
    "release": "",           # release the sprint was found for, e.g. '26.08'
    "user": "",              # display name of whoever is signed in
}


# --------------------------------------------------------------------------- #
#  Helpers
# --------------------------------------------------------------------------- #
# Reading a project's create metadata is the slowest call in the app, and it
# is the same answer for every ticket in that project - so it is cached by
# project, and one lock stops two tickets racing to fetch the same one.
_CTX_LOCK = threading.Lock()


def _project_of(parent_key: str) -> str:
    """The project a ticket belongs to. One cheap call, then cached."""
    known = STATE["projects"].get(parent_key)
    if known:
        return known
    project = STATE["client"].get_parent_context(parent_key)["project_key"]
    STATE["projects"][parent_key] = project
    return project


def _project_context(project: str):
    """(subtask_type_id, create_fields) for a project. The expensive one."""
    have = STATE["contexts"].get(project)
    if have:
        return have
    with _CTX_LOCK:
        have = STATE["contexts"].get(project)      # another thread may have won
        if have:
            return have
        ctx = STATE["client"].get_project_subtask_context(project)
        STATE["contexts"][project] = ctx
        return ctx


def _context_for(parent_key: str):
    """(project_key, subtask_type_id, create_fields) for a parent issue."""
    project = _project_of(parent_key)
    subtask_id, create_fields = _project_context(project)
    return project, subtask_id, create_fields


def _warm_project(project: str) -> None:
    """Fetch a project's create metadata in the background, so step 4 is instant."""
    try:
        _project_context(project)
    except Exception:  # noqa: BLE001 - a warm-up that fails costs nothing
        pass


def build_plan(parent, task_value, title, start, end, mode):
    """Day-by-day plan, using the same rules as the CLI."""
    return core.build_plan(
        start=start, end=end, mode=mode, attendance=STATE["attendance"],
        parent=parent, leave_parent=STATE["leave_parent"],
        summary_base=f"{task_value}-{title}" if title else task_value,
        task_value=task_value, logged=STATE["logged"],
    )


def _hours(minutes: int) -> str:
    """'8h 30m' for a count of minutes, '' for nothing at all."""
    return core.format_minutes(minutes) if minutes else ""


def _remember_logged(plan, results) -> None:
    """
    Fold the work logs we have just written into the timesheet we hold.

    A sub-task now covers several days, so this credits the days themselves -
    only the ones this run actually wrote. Without it a second round in the
    same session would offer those days again: the timesheet moved on, but our
    copy of it did not.
    """
    for entry, row in zip(plan, results):
        parent = entry.get("parent") or ""
        for written in row.get("fresh") or []:
            mins = int(written.get("minutes") or 0)
            if mins <= 0:
                continue
            day = STATE["logged"].setdefault(
                written["date"], {"minutes": 0, "parents": {}, "issues": {}})
            day["minutes"] += mins
            if parent:
                day["parents"][parent] = day["parents"].get(parent, 0) + mins


# --------------------------------------------------------------------------- #
#  Routes
# --------------------------------------------------------------------------- #
@app.route("/")
def index():
    folder, name = _find_index()
    return send_from_directory(folder, name)


@app.get("/api/session")
def api_session():
    """Who is signed in, so a reloaded page doesn't look logged out."""
    return jsonify(signed_in=bool(STATE["client"]), user=STATE["user"],
                   release=STATE["release"],
                   attendance_days=len(STATE["attendance"]),
                   logged_days=len(STATE["logged"]))


def _body() -> dict:
    """
    The JSON a request carried, or {} when it carried none.

    `request.json` *raises* on a POST with no JSON content type - which is
    exactly what a bodyless `fetch(url, {method:"POST"})` sends. Reading it
    directly meant Sign out's reset never ran and the session survived it.
    """
    return request.get_json(silent=True) or {}


@app.post("/api/reset")
def api_reset():
    """
    Clear the work gathered so far: sprint, ticket, attendance, release.

    With `keep_session` the sign-in survives — that is what a page reload wants,
    a clean form but no need to sign in again. Without it the session goes too,
    which is what Start over wants.

    Either way the saved password stays put, and so does the attendance
    portal's token: that is a cache, not something the user entered.
    """
    keep = bool(_body().get("keep_session"))
    STATE.update(sprint_issues=[], leave_parent=None, attendance={},
                 contexts={}, projects={}, release="", logged={}, logged_range="")
    if not keep:
        STATE.update(client=None, user="")
    return jsonify(ok=True, signed_in=bool(STATE["client"]))


@app.post("/api/login")
def api_login():
    data = _body()
    base_url = (data.get("base_url") or "https://tracking.i2cinc.com/").strip()
    method = data.get("method", "password")
    no_verify = data.get("no_verify_ssl", False)
    try:
        if method == "cookie":
            cookie = (data.get("cookie") or "").strip()
            if not cookie:
                return jsonify(ok=False, error="Paste your Jira session cookie first."), 400
            client = core.JiraClient(base_url, session_cookie=cookie,
                                     verify_ssl=not no_verify)
        else:
            username = (data.get("username") or "").strip()
            password = data.get("password") or ""
            if not password:
                return jsonify(ok=False, error="Password / token is required."), 400
            client = core.JiraClient(base_url, username, password,
                                     verify_ssl=not no_verify)
        who = client.verify_login()
    except core.LoginError as exc:
        return jsonify(ok=False, error=str(exc)), 401
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=f"Unexpected error: {exc}"), 500
    STATE.update(client=client, contexts={}, projects={}, user=who)
    # Remember a working password/PAT so the one-click button needs no typing.
    saved_to = ""
    if method != "cookie" and data.get("remember", True):
        saved_to = jira_credentials.save(base_url, username, password)
    return jsonify(ok=True, user=who, saved=saved_to)


@app.post("/api/sso/start")
def api_sso_start():
    """Begin the automatic Microsoft sign-in; the UI polls /api/sso/status."""
    data = _body()
    base_url = (data.get("base_url") or "https://tracking.i2cinc.com/").strip()
    verify_ssl = not data.get("no_verify_ssl", False)

    def _adopt(client, who):
        STATE.update(client=client, contexts={}, projects={}, user=who)

    ok, err = sso_login.MANAGER.start(base_url, verify_ssl, _adopt,
                                      interactive=bool(data.get("interactive")))
    if not ok:
        return jsonify(ok=False, error=err), 400
    return jsonify(ok=True)


@app.get("/api/sso/status")
def api_sso_status():
    return jsonify(sso_login.MANAGER.status())


@app.post("/api/sso/cancel")
def api_sso_cancel():
    sso_login.MANAGER.cancel()
    return jsonify(ok=True)


@app.post("/api/sprint")
def api_sprint():
    if not STATE["client"]:
        return jsonify(ok=False, error="Not logged in."), 401
    data = _body()
    client = STATE["client"]
    st = (data.get("st") or "").strip()
    release = (data.get("release") or "").strip()
    sprint = (data.get("sprint") or "").strip()
    # Board id is optional: blank means "work it out from the sprint name".
    board_raw = str(data.get("board") or "").strip()
    board = int(board_raw) if board_raw.isdigit() else None
    others: list = []
    try:
        if st and release:
            # Team + release: match the sprint however that team spells it.
            found = client.find_release_sprint(st, release)
            sprint = found["name"]
            meta = {"sprint_id": found["id"], "sprint_name": found["name"],
                    "board_id": found["board_id"]}
            others = found["others"]
            issues = client.issues_in_sprint(found["id"])
        elif sprint:
            issues, meta = client.search_sprint(sprint, board_id=board)
        else:
            return jsonify(ok=False,
                           error="Enter your ST and the release (e.g. ST19 and 26.08)."), 400
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=str(exc)), 400
    if not issues:
        return jsonify(ok=False, error=f"No issues found in sprint '{sprint}'."), 404
    STATE["sprint_issues"] = issues
    STATE["release"] = release or STATE["release"]
    # auto-detect the Planned Leaves parent
    STATE["leave_parent"] = None
    for k, s, *rest in issues:
        if "planned leave" in s.lower():
            STATE["leave_parent"] = k
            break
    # Remember which project each ticket is in, and start reading those
    # projects' create metadata now: it is what step 4 waits on when a ticket
    # is picked, and by then it is usually already here.
    for k, _s, *rest in issues:
        if rest and rest[0]:
            STATE["projects"][k] = rest[0]
    for project in list(dict.fromkeys(STATE["projects"].values()))[:3]:
        threading.Thread(target=_warm_project, args=(project,), daemon=True).start()
    return jsonify(ok=True,
                   issues=[{"key": k, "summary": s,
                            "project": (rest[0] if rest else "")}
                           for k, s, *rest in issues],
                   leave_parent=STATE["leave_parent"],
                   board_id=meta.get("board_id"),
                   sprint_name=meta.get("sprint_name"),
                   others=others)


@app.post("/api/tasks")
def api_tasks():
    if not STATE["client"]:
        return jsonify(ok=False, error="Not logged in."), 401
    data = _body()
    parent = (data.get("parent") or "").strip()
    # The page knows each ticket's project from step 2, so the usual case needs
    # no lookup at all: the Task list is per project and already cached.
    project = (data.get("project") or "").strip() or STATE["projects"].get(parent, "")
    try:
        if not project:
            project = _project_of(parent)
        _stid, create_fields = _project_context(project)
        _fid, opts = STATE["client"].get_task_options(create_fields)
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=str(exc)), 400
    return jsonify(ok=True, tasks=[v for v, _ in opts], project=project)


@app.post("/api/attendance")
def api_attendance():
    if "file" not in request.files:
        return jsonify(ok=False, error="No file uploaded."), 400
    f = request.files["file"]
    if not f.filename.lower().endswith((".xlsx", ".xls", ".csv")):
        return jsonify(ok=False, error="Please upload a .xlsx / .xls / .csv file."), 400
    suffix = os.path.splitext(f.filename)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    try:
        f.save(tmp.name)
        tmp.close()
        attendance = core.load_attendance(tmp.name)
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=f"Could not read attendance: {exc}"), 400
    finally:
        try:
            os.unlink(tmp.name)
        except OSError:
            pass
    if not attendance:
        return jsonify(ok=False, error="No usable rows found in the file."), 400
    STATE["attendance"] = attendance
    days = sum(1 for r in attendance.values() if r["category"] == "work")
    leaves = sum(1 for r in attendance.values() if r["category"] == "leave")
    return jsonify(ok=True, work_days=days, leave_days=leaves, total_rows=len(attendance))


@app.post("/api/attendance/fetch")
def api_attendance_fetch():
    """Pull the sheet off the attendance portal for the release from step 2."""
    d = _body()
    release = (d.get("release") or STATE["release"] or "").strip()
    start = (d.get("start") or "").strip()
    end = (d.get("end") or "").strip()
    if not release and not (start and end):
        return jsonify(ok=False,
                       error="Find the sprint in step 2 first — the release "
                             "tells the portal which range to read."), 400

    def _adopt(attendance):
        STATE["attendance"] = attendance

    ok, err = attendance_portal.MANAGER.start(
        release, start, end, _adopt, interactive=bool(d.get("interactive")))
    if not ok:
        return jsonify(ok=False, error=err), 400
    return jsonify(ok=True)


@app.get("/api/attendance/status")
def api_attendance_status():
    return jsonify(attendance_portal.MANAGER.status())


@app.post("/api/attendance/cancel")
def api_attendance_cancel():
    attendance_portal.MANAGER.cancel()
    return jsonify(ok=True)


@app.post("/api/logged")
def api_logged():
    """
    Read your own Jira worklogs for the range, so the app can tell a day that
    is already done from one that is missing hours or is short of them.

    A failure here is not fatal: the answer says so, nothing is subtracted, and
    the run behaves as it did before the check existed.
    """
    if not STATE["client"]:
        return jsonify(ok=False, error="Not logged in."), 401
    d = _body()
    start, end = (d.get("start") or "").strip(), (d.get("end") or "").strip()
    if not (start and end):
        return jsonify(ok=False, error="Pick a start and end date first."), 400
    key = f"{start}|{end}"
    if d.get("force") or STATE["logged_range"] != key:
        try:
            STATE["logged"] = STATE["client"].my_worklogs(start, end)
            STATE["logged_range"] = key
        except Exception as exc:  # noqa: BLE001
            STATE.update(logged={}, logged_range="")
            return jsonify(ok=False, error=str(exc)), 400
    logged = STATE["logged"]
    return jsonify(ok=True, days=len(logged),
                   minutes=sum(int(r["minutes"]) for r in logged.values()),
                   start=start, end=end)


@app.post("/api/days")
def api_days():
    """
    One row per date in the range: the hours that day is worth and whether it
    is work, leave or nothing at all. Step 4's calendar is drawn from this, so
    the days you can put a ticket on are exactly the days that will be created.
    """
    d = _body()
    start, end = (d.get("start") or "").strip(), (d.get("end") or "").strip()
    mode = d.get("mode", "attendance")
    if not (start and end):
        return jsonify(ok=False, error="Pick a start and end date first."), 400
    try:
        dates = core.daterange(start, end, True)
    except ValueError as exc:
        return jsonify(ok=False, error=str(exc)), 400
    days = []
    for iso in dates:
        day = core.remaining_day(iso, mode, STATE["attendance"],
                                 logged=STATE["logged"],
                                 leave_parent=STATE["leave_parent"])
        rec = STATE["attendance"].get(iso) or {}
        worth = day["target_work"] + sum(m for _, m in day["target_leaves"])
        left = day["work"] + sum(m for _, m in day["leaves"])
        days.append({
            "date": iso,
            # `minutes` is what is still to be logged, so everything built on
            # it - the calendar, the splits, the counters - is about the hours
            # that are actually missing rather than the whole day again.
            "minutes": day["work"],
            "hours": _hours(day["work"]),
            "target": day["target_work"],
            "target_hours": _hours(day["target_work"]),
            "logged": day["logged_work"],
            "logged_hours": _hours(day["logged_work"]),
            "day_logged": day["logged"],
            "day_logged_hours": _hours(day["logged"]),
            # The whole day, work plus any leave top-up, so the UI can show
            # what Jira has against what the day is worth in this mode.
            "worth": worth,
            "worth_hours": _hours(worth),
            "left": left,
            "left_hours": _hours(left),
            "full": bool(worth and not left),
            "category": rec.get("category") or "none",
            "leaves": [{"title": t, "hours": core.format_minutes(m)}
                       for t, m in day["leaves"]],
        })
    return jsonify(ok=True, days=days, leave_parent=STATE["leave_parent"],
                   logged_days=len(STATE["logged"]),
                   logged_checked=bool(STATE["logged_range"]))


@app.post("/api/plan")
def api_plan():
    if not STATE["client"]:
        return jsonify(ok=False, error="Not logged in."), 401
    d = _body()
    mode = d.get("mode", "attendance")
    # Both modes read the sheet; they differ only in what a full day logs.
    if mode in ("attendance", "static") and not STATE["attendance"]:
        return jsonify(ok=False, error="Get the attendance sheet first."), 400
    # Two shapes: `items` is the per-date picture from step 4 (one or more
    # tickets a day, split by percentage); `parent` is the older one-ticket-
    # for-the-whole-range call, still used by the CLI-shaped request.
    items = d.get("items") or None
    unassigned: list = []
    already: list = []
    try:
        if items:
            plan, unassigned, already = core.build_plan_items(
                start=d["start"], end=d["end"], mode=mode,
                attendance=STATE["attendance"], items=items,
                leave_parent=STATE["leave_parent"], logged=STATE["logged"])
        else:
            plan = build_plan(d["parent"], d["task_value"], d.get("title", ""),
                              d["start"], d["end"], mode)
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=str(exc)), 400
    if any(e["kind"] == "leave" for e in plan) and not STATE["leave_parent"]:
        return jsonify(ok=False, error="Leave days found but no 'Planned Leaves' "
                       "issue in the sprint. Can't route them."), 400
    # One sub-task per ticket+Task, holding whole days up to the cap, each with
    # its own day-by-day work logs. The day list above is what it was built
    # from, and is still what the calendar and the counts are drawn from.
    days = len({e["date"] for e in plan})
    packed = core.pack_plan(plan)
    return jsonify(ok=True, plan=packed, unassigned=unassigned, already=already,
                   day_count=days, cap=core.format_minutes(core.PACK_CAP))


@app.post("/api/run")
def api_run():
    if not STATE["client"]:
        return jsonify(ok=False, error="Not logged in."), 401
    client = STATE["client"]
    d = _body()
    plan = d.get("plan", [])
    status_path = ["In Progress", "Done"]

    # Resolve each parent's create metadata once, up front. It is the slowest
    # call in the whole run and it is shared, so doing it here keeps the
    # workers below from racing over the cache — or repeating the work.
    parents = sorted({e["parent"] for e in plan if e.get("parent")})
    try:
        if len(parents) > 1:
            with ThreadPoolExecutor(max_workers=len(parents)) as pool:
                list(pool.map(_context_for, parents))
        elif parents:
            _context_for(parents[0])
    except Exception as exc:  # noqa: BLE001
        return jsonify(ok=False, error=f"Couldn't read the issue setup: {exc}"), 400

    def create_one(entry: dict) -> dict:
        """
        One sub-task: create it, log each of its days, walk it to Done.

        A retry hands back the row it got last time under `_row`, so a run that
        died half way carries on from where it stopped: the issue is not
        created twice, and the days already logged are not logged again -
        `logged_dates` remembers exactly which of them went in.
        """
        prior = entry.get("_row") or {}
        logs = core.plan_logs(entry)
        row = {"date": entry["date"], "kind": entry["kind"], "ok": False,
               "key": prior.get("key"),
               "summary": core.entry_summary(entry),
               "days": len(logs),
               "logged_dates": list(prior.get("logged_dates") or []),
               "fresh": [],                    # what this call put on Jira
               "worklogged": bool(prior.get("worklogged")),
               "done": bool(prior.get("done")), "error": None}
        try:
            project, subtask_id, create_fields = _context_for(entry["parent"])
            if not row["key"]:
                fields = {}
                if client.assignee_ref:
                    fields["assignee"] = client.assignee_ref
                if entry["estimate"]:
                    fields["timetracking"] = {"originalEstimate": entry["estimate"]}
                if entry["task_value"]:
                    fid, ref = client.resolve_task_field(create_fields,
                                                        entry["task_value"])
                    if fid and ref:
                        fields[fid] = ref
                row["key"] = client.create_subtask(
                    project, entry["parent"], subtask_id, row["summary"],
                    extra_fields=fields)
            # The days of one sub-task are separate work logs against the same
            # issue, so they go out together. Each answers for itself: a day
            # that lands is remembered even when another fails, which is what
            # lets a retry ask for only what is still missing.
            comment = core.entry_base(entry)
            pending = [l for l in logs if l["date"] not in row["logged_dates"]]

            def write(log):
                try:
                    client.log_work(row["key"], log["date"], log["hours"],
                                    comment=comment)
                    return log, None
                except Exception as exc:  # noqa: BLE001
                    return log, str(exc)

            refused = []
            for log, failure in core.in_parallel(write, pending, limit=LOG_WORKERS):
                if failure:
                    refused.append(failure)
                    continue
                row["logged_dates"].append(log["date"])
                # Only time written *here* is news to our copy of the
                # timesheet; a retry must not count last round's worklog twice.
                row["fresh"].append({"date": log["date"],
                                     "minutes": core.duration_minutes(log["hours"])})
            row["logged_dates"].sort()
            row["worklogged"] = not refused
            if refused:
                raise RuntimeError(refused[0])
            if not row["done"]:
                client.move_through(row["key"], status_path,
                                    task_value=entry["task_value"] or "",
                                    check_all=True, project=project)
            row["done"] = True
            row["ok"] = True
        except Exception as exc:  # noqa: BLE001
            row["error"] = str(exc)
        return row

    # Each day is its own sub-task and independent of the others, so they go
    # out together instead of one-at-a-time. Results keep the plan's order.
    workers = max(1, min(RUN_WORKERS, len(plan)))
    if workers > 1:
        with ThreadPoolExecutor(max_workers=workers) as pool:
            results = list(pool.map(create_one, plan))
    else:
        results = [create_one(e) for e in plan]

    ok = sum(1 for r in results if r["ok"])
    # Whatever we managed to log is on the timesheet now, so treat it as read:
    # a second round in this session sees those days as done.
    _remember_logged(plan, results)
    return jsonify(ok=True, results=results, summary=f"{ok}/{len(results)} done")


def _open_browser():
    webbrowser.open("http://127.0.0.1:5000")


if __name__ == "__main__":
    print("Jira Timesheet UI running at  http://127.0.0.1:5000")
    threading.Timer(1.0, _open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
