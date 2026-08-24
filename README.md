<div align="center">

# 🗓️ Jira Task Logging

**Fill in a whole release worth of Jira time in one pass.**

It reads your attendance, reads what your Jira timesheet already has, works out
what each day is still owed, groups those days into sub-tasks, logs a work log
per day, and walks each sub-task to **Done**.

[![Python](https://img.shields.io/badge/Python-3.10%2B-3776AB?logo=python&logoColor=white)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-web%20UI-000000?logo=flask&logoColor=white)](https://flask.palletsprojects.com/)
[![Selenium](https://img.shields.io/badge/Selenium-headless%20SSO-43B02A?logo=selenium&logoColor=white)](https://www.selenium.dev/)
[![Jira](https://img.shields.io/badge/Jira-Server%20%2F%20DC-0052CC?logo=jira&logoColor=white)](https://www.atlassian.com/software/jira)
[![Runs locally](https://img.shields.io/badge/Runs-100%25%20locally-success)](#)

</div>

---

## ✨ Why

Logging a release by hand is twenty-odd sub-tasks, a work log per day, and a
walk to **Done** on each one. This does the whole thing in two clicks — and
never logs the same hour twice.

| | |
| --- | --- |
| 🔍 **Reads your attendance** | straight from the portal's API — no copy-paste, no spreadsheet |
| 🧮 **Works out what each day is worth** | floors the hours, credits half hours, adds **Short Leave** for a light day |
| 🧾 **Subtracts what Jira already has** | a half-finished range picks up exactly the days that are missing or short |
| 📦 **Packs days into sub-tasks** | a release becomes a handful of sub-tasks, not twenty — one work log per day, always |
| 🚀 **Runs them in parallel** | overlapping the waiting turns a 5.1s run into 2.1s |
| 🔐 **Keeps your sign-in safe** | Windows Credential Manager, never a plaintext file |

```
                 attendance portal ──┐
                                     ├──► what each day is worth ──┐
your Jira worklogs ──────────────────┘                            ├──► sub-tasks
                 you: which ticket, which Task, a title ──────────┘
```

There are two ways in — a local web page (the usual one) and a command line
tool. Both run on your machine and talk only to your Jira and your attendance
portal.

---

## ⚡ Quick start

```powershell
pip install requests openpyxl flask selenium keyring browser_cookie3 tzdata
python app.py
```

Then open **<http://127.0.0.1:5000>**, and do the first run small — narrow the
dates to two or three days and press **Preview plan** before anything is
written. Full walk-through in [§1.6 First run](#16-first-run).

Prefer a terminal? `python jira_logging_utility.py --dry-run` previews the same
plan and creates nothing.

---

## 📖 Contents

| | Section | What's in it |
| --- | --- | --- |
| **1** | [Configure it](#1-configure-it) | requirements, install, servers, where your sign-in is kept, tunables, first run |
| **2** | [The web UI, step by step](#2-the-web-ui-step-by-step) | sign in → sprint → dates → work items → review → results |
| **3** | [What a day is worth](#3-what-a-day-is-worth) | the hour rules, both modes, worked examples |
| **4** | [What Jira already has](#4-what-jira-already-has) | the timesheet read-back, and why hours can't go in twice |
| **5** | [How sub-tasks are built](#5-how-sub-tasks-are-built) | grouping, packing, naming, and why a run is quick |
| **6** | [The command line](#6-the-command-line) | every flag, and a fully non-interactive run |
| **7** | [How it is put together](#7-how-it-is-put-together) | the files, the shared rules, the API endpoints |
| **8** | [When something goes wrong](#8-when-something-goes-wrong) | every error message, and what to do about it |

> [!TIP]
> New here? Read [§1.6 First run](#16-first-run), then
> [§3 What a day is worth](#3-what-a-day-is-worth) — between them they cover
> everything the tool will actually do to your Jira.

> [!IMPORTANT]
> Nothing leaves your machine. Every call goes to your own Jira and your own
> attendance portal, and your password lives in Windows Credential Manager.

---

## 1. Configure it

### 1.1 What you need

| | |
| --- | --- |
| **Python 3.10+** | `python --version` |
| **Microsoft Edge or Chrome** | only for the automatic sign-in and the attendance fetch |
| **Network** | reachable Jira (VPN if your Jira needs one) |

### 1.2 Install

Everything, web UI included:

```powershell
pip install requests openpyxl flask selenium keyring browser_cookie3 tzdata
```

The command line tool alone needs far less:

```powershell
pip install requests openpyxl
```

What each package is for, so you can leave out what you don't want:

| Package | Needed for | Without it |
| --- | --- | --- |
| `requests` | every Jira call | nothing works |
| `openpyxl` | reading an `.xlsx` attendance export | export as `.csv` instead |
| `flask` | the web page | command line only |
| `selenium` | driving a headless browser: the attendance fetch and Microsoft sign-in | sign in with a password/PAT and drop the export on the page |
| `keyring` | keeping your password in Windows Credential Manager | falls back to a base64 file — see 1.4 |
| `browser_cookie3` | borrowing a Jira session from a browser you are already signed in to | one less sign-in shortcut |
| `tzdata` | resolving `Asia/Karachi` and friends on Windows | falls back to this machine's timezone, which is normally the same answer |

### 1.3 Point it at your servers

Defaults live at the top of each module — change them there, or pass a flag.

| Setting | Where | Default |
| --- | --- | --- |
| Jira base URL | the **Jira URL** box on the page, or `--base-url` | `https://tracking.i2cinc.com/` |
| Attendance portal | `attendance_portal.py` › `PORTAL_URL` | `https://attendance.i2cinc.com/employee/attendance` |
| Attendance API | `attendance_portal.py` › `API_BASE` | `https://attendance-server-pilot.i2cinc.com/api/v1` |
| Jiras with no Microsoft SSO | `sso_login.py` › `PASSWORD_ONLY_HOSTS` | `("tracking.i2cinc.com",)` |
| Web UI port | `app.py` › `app.run(... port=5000)` | `5000` |

### 1.4 Where your sign-in is kept

- **Windows Credential Manager**, under the service name `jira-timesheet`, one
  entry per Jira host. This is where your password or PAT goes when you tick
  *remember* — which is what makes the one-click button work later.
- If `keyring` is missing it falls back to `.jira-credentials` next to the
  scripts. That file is **base64, not encryption** — keep `keyring` installed.
  It is gitignored either way.
- `.sso-browser-profile/` is the browser profile the automation drives. It
  holds the live Microsoft, Jira and portal sessions, which is what lets the
  fetch run with nothing on screen. It is gitignored, it can grow to a few
  hundred MB, and deleting it simply means signing in once more.

### 1.5 Behaviour you can tune

| Knob | Where | Default | Means |
| --- | --- | --- | --- |
| `STANDARD_DAY` | `jira_logging_utility.py` | `8 * 60` | a full logged day |
| `FULL_DAY` | `jira_logging_utility.py` | `9 * 60` | attendance at or above this is a long day |
| `PACK_CAP` | `jira_logging_utility.py` | `24 * 60` | most one sub-task may hold (`--pack-cap` on the CLI) |
| `POOL_SIZE` | `jira_logging_utility.py` | `32` | sockets kept open to Jira |
| `RUN_WORKERS` | `app.py` | `8` | sub-tasks on the go at once |
| `LOG_WORKERS` | `app.py` | `4` | days of one sub-task logged at once |
| status path | `--status-path` | `In Progress,Done` | the statuses each sub-task is walked through |

`RUN_WORKERS × LOG_WORKERS` is the most requests in flight at once, so keep
`POOL_SIZE` at least that big. Turn them down if your Jira starts answering
with 429s.

### 1.6 First run

```powershell
python app.py          # then open http://127.0.0.1:5000
```

Do the first one small and see the plan before it writes anything:

1. Sign in with your username and password (or a PAT with the username blank).
   Let it save; from then on **Sign in with Microsoft** needs no typing.
2. Enter your **ST** and **release** — e.g. `ST12` and `26.08`.
3. Let the attendance fetch run. First time it is ~20s, after that ~2s.
4. Narrow the dates to two or three days.
5. Assign a day, press **Preview plan**, and read it: the sub-tasks, their
   days, the hours, the names. Nothing has been created yet.
6. **Create & log all**, then check one sub-task in Jira.

For the command line, the same caution is `--dry-run`:

```powershell
python jira_logging_utility.py --dry-run
```

---

## 2. The web UI, step by step

### Step 1 — Sign in

Two routes to the same place:

- **Username + password / PAT.** Straight through the REST API, no browser. On
  Jira Server, if password login is blocked, create a Personal Access Token,
  leave the username blank, and paste the token as the password.
- **Sign in with Microsoft.** Everything happens in the backend, in this order:
  the password you saved earlier (one REST call, effectively instant); a live
  Jira session from a browser on this machine (`browser_cookie3`); a headless
  SSO walk on the app's own profile. It gives up in about 10s rather than
  hanging, and says which of those failed.

`tracking.i2cinc.com` has no Microsoft SSO — its login page is the stock
Atlassian form — so there the button is really "use my saved password".

A page reload keeps you signed in; **Start over** clears the session.

### Step 2 — Sprint & issue

Enter the **ST** (`ST19`, `ST-19` or `19`) and the **release** (`26.8` or
`26.08`, treated as the same). The match is made on the team number and the
release number, not the text, so one lookup handles `ST12-26.8`,
`ST-19 - 26.08 Release`, `26.08 - ST4 Sprint 8` and `ST 7 26.08` alike.

If a team runs several sprints in a release, the active one wins, then the
highest sprint number; the rest are listed. If nothing matches, the error names
the sprints that *do* exist for that ST and unlocks the Sprint / Board id boxes
so you can type the name yourself.

The sprint's tickets are loaded, the **Planned Leaves** parent is detected from
the issue whose summary contains "Planned Leaves", and the Task metadata for
those tickets' project starts loading in the background — which is why the Task
dropdown in step 4 opens instantly rather than after a wait.

### Step 3 — Dates & attendance

**The attendance sheet fetches itself.** The portal is a single-page app over a
REST API, so `attendance_portal.py` talks to that API rather than reading the
screen: a headless browser opens the portal only long enough to mint a session
and hand over the bearer token, then two requests do the rest.

| Call | Gives us |
| --- | --- |
| `GET /admin/releases/…/withExceptionCutOff` | `26.08` → `2026-07-22 … 2026-08-23` |
| `GET /employee/attendances/getAttendanceAndSummary?startDate&endDate` | one record per day |

Each record carries `totalHours` as a number plus explicit `holiday` / `leave` /
`dayOff` / `absent` flags — more reliable than parsing `9h 26m` out of a grid.
The leave *type* is in `remarks` (`Casual`) and the *approval state* in
`leaveStatus` (`Pending`), so a pending Casual Leave still logs as
`Casual Leave`. Absent days and the day still in progress log nothing.

A cold fetch is about **20s**, mostly the portal's own boot and the Microsoft
round trip. The token is then held while it is valid, so later fetches are
about **2s** with no browser at all. If the portal session has lapsed, the
headless browser presses **Sign in with Microsoft** itself. If the fetch cannot
be done here at all, the page falls back to dropping the `.xlsx` on it.

**The dates** are pre-filled from the release and never run past today. Change
them freely; everything below follows.

**How hours are set** — two modes, differing *only* in what a full weekday
logs. See section 3.

**Already on your Jira timesheet** — see section 4.

### Step 4 — Work items

A calendar of the range, drawn from the attendance itself, so what you can
click is exactly what will be created.

| On the calendar | Means |
| --- | --- |
| outlined, hours under it | a day to click — the hours shown are the hours still **missing** |
| amber outline | partly logged in Jira; hovering says `4h of 10h already in Jira` |
| a green ✓, no number | already logged in full — nothing to choose, nothing to click |
| purple | leave or public holiday — it goes to Planned Leaves on its own |
| rose, with a `!` | **absent** — the portal recorded no hours and no leave. Nothing is logged for it automatically and **Select unassigned** skips it; click it and type what it is worth |
| green fill | already assigned |
| blank / dim | not in the range, or nothing worked |

Under it, one line keeps score — *"**4 days** still need hours (4h 2m) · 20
done in Jira"*.

Click a day, shift-click for a run, ctrl-click to pick out odd ones. Then give
them a **ticket**, a **Task** and — optionally — a **title** for what you were
doing. Press **Add to plan**. A whole sprint on one ticket is two clicks:
**Select unassigned**, then Add.

**More than one ticket in a day.** Press **+ Add ticket** and each row takes a
share of the day. The same mechanism covers two Tasks on one ticket — pick the
same ticket on both rows.

| Divide by | How |
| --- | --- |
| **Percent** | drag the sliders. An 8h day at 60/40 becomes `4h 48m` + `3h 12m`. Must come to 100% |
| **Hours** | type them: `5h 30m` + `2h 30m`. Must come to the day's hours |

Either way the leftover is spelled out — *"2h of 8h still to give out"*, *"that
is 30m more than the day's 8h"* — and **Add to plan** stays off until it
balances. **Split evenly** fixes it in whichever unit you are using, and
switching between the two carries your division across. The hours box takes
`5h 30m`, `5.5`, `5:30` or `330m`.

Not a minute is lost or invented: the hours are divided into whole minutes,
largest remainder first, and the parts always add back up to the day. Typing
`5h 30m` gets you `5h 30m` in Jira.

Typing absolute hours needs every selected day to be worth the same, so a run
of 8h days can be typed at once; mixing an 8h day with a 7h one keeps you on
percentages, and says so.

Days you never assign are simply skipped, and the counter says so
(`14 of 19 work days`) rather than letting you find out afterwards.

### Step 5 — Review & run

One line per sub-task, grouped into a panel per **ticket + Task**:

```
 10 sub-tasks · 160h · 20 work logs · 8 work, 2 leave                 [Show the days]

 ST12-5876  Invoice module – billing engine        DEVELOPMENT   40h · 2 sub-tasks   same title for all
        HOURS   WHICH DAYS                    TITLE                NAME IN JIRA
   >    24h     Mon, Jul 27 – Wed, Jul 29 · 3 days   [ Invoice totals  v ]  Development-Invoice totals
   >    16h     Thu, Jul 30 – Fri, Jul 31 · 2 days   [ No title        v ]  Development
```

The top bar **stays with you as you scroll**, so the totals are in view however
long the plan is. **Create & log all** sits at the foot of the step, where it
always has. A month across four tickets is about 900px of page instead of the
several thousand a card each used to take.

Each row folds out — the `>` on the row, or **Show the days** for all of them at
once — to show its day chips (`Mon 27 Jul 8h`) and, where the days were titled
differently, which titles it could go by. A row opens by itself when there is
something to see: mixed titles, or a failure.

Under **Preview plan** in step 4, whatever the plan left out is a count you can
read at a glance, with the days as chips — and a long list folds away:

> !  **2 working days with no ticket** — nothing will be created for them
> `Mon, Jul 27` `Wed, Jul 29`
>
> ok **18 days already on your Jira timesheet** — left untouched · *show the days*

### Step 6 — Results

The same cards, now carrying `done` or `failed`, the new Jira key, and how far
each got — a half-written one says so (*"failed · ST12-7003 · 1 of 3 logged"*)
with the reason underneath.

**Retry failed** re-runs only the failures and each resumes exactly where it
stopped: the sub-task is not created twice, and only the days still missing are
logged. **Log another task** keeps the sign-in, sprint, dates and attendance,
and clears just the work items.

---

## 3. What a day is worth

Total Hours are floored to the whole hour, with a half hour credited when the
minutes are **50 or more** (`9h 45m` → `9h`, `9h 58m` → `9h 30m`).

The two modes differ **only** in what a full weekday logs. Short days,
weekends, leave and public holidays behave identically in both.

| Day type | Logged | Created under |
| --- | --- | --- |
| Weekday, floored ≥ 9h | floored − 1h *(Static 8h: a flat 8h)* | your ticket |
| Weekday, floored < 9h | floored | your ticket **+** a **Short Leave** sub-task for (8h − floored) under Planned Leaves |
| Weekend with hours | floored, in full | your ticket |
| Public holiday / leave | 8h | Planned Leaves |
| Absent | nothing automatically — *what you type, if you claim the day* | your ticket |
| Today | nothing | — |

Examples:

- `9h 13m` → **8h**
- `10h 50m` → **9h 30m**
- `7h 2m` → **7h** in the ticket **+ 1h Short Leave**
- `6h` → **6h** in the ticket **+ 2h Short Leave**

**Absent days are yours to decide.** When the portal marks a day absent — or
marks it anything that is not leave on a working day — there are no hours to go
by, and filing it as leave would be a guess. So nothing is created for it on its
own: it shows on the calendar in rose with a `!`, **Select unassigned** passes
over it, and the day counter names how many are still unclaimed. Click one (or a
run of them), pick a ticket and a Task, and type the hours you want against it —
that is the only thing the day is worth. Whatever Jira already holds for that
date still comes off the top, so claiming an absent day cannot log the same hour
twice. Because they have no hours of their own, absent days are assigned on
their own rather than mixed into a selection of ordinary days.

The estimate on each sub-task is whatever it holds in total.

---

## 4. What Jira already has

Before anything is planned, your own worklogs for the range are read back
(`worklogAuthor = currentUser()`), and every day is trimmed by what it already
has:

| The day in Jira | Result |
| --- | --- |
| nothing logged | the whole day is offered |
| logged for **less** than section 3 says | only the **difference** (4h of a 10h day → a 6h sub-task) |
| logged in **full** | left completely alone — no sub-task, nothing to click |
| logged for **more** than the day is worth | left alone |

So the same hours can never go in twice: re-running over a half-finished range
picks up exactly the days that are missing or short. Work and Planned Leaves
are counted separately, so a short day whose `6h` *and* whose `2h` Short Leave
are both in Jira counts as done.

**Which day a worklog belongs to.** Jira answers with the *instance's* clock,
not yours: 9am in Karachi comes back as `2026-08-10T21:00:00.000-0700`, so
reading the date off the front of that string would file Tuesday's work under
Monday. Every worklog is converted into your own timezone (your Jira profile's
`timeZone`, or this machine's if the OS has no zone database) before it is
dated, and the search runs a day wide at each end because Jira matches
`worklogDate` in its own timezone too. The per-day totals then agree exactly
with **Issues › Timesheet – Worklog Tracker**.

**Day by day** on the timesheet strip opens the figures so they can be read
against that page line by line:

| Date | In Jira now | Day is worth | Still to log |
| --- | --- | --- | --- |
| Mon, Jul 27 | `8h` | `9h` | `1h` |
| Tue, Aug 11 | `11h 11m` | `11h` | done |
| Wed, Aug 19 | nothing | `8h` | `8h` |

**Day is worth** follows the mode you picked, so switching *From attendance* /
*Static 8h* redraws it. **Re-check Jira** reads the timesheet again, for when
you have logged something in Jira itself since. If Jira refuses the search it
says so and offers every day as usual — the check never blocks a run, but do
glance at the Worklog Tracker in that case.

---

## 5. How sub-tasks are built

Days are grouped by **ticket + Task** and packed into sub-tasks of up to
**24h**, so a release becomes a handful of sub-tasks instead of twenty-odd:

```
ST12-5876 · Development   42h → 2 sub-tasks   24h  (22nd 8h · 23rd 8h · 24th 8h)
                                              18h  (27th 8h · 28th 5h 42m · 29th 4h 18m)
ST12-5901 · Testing       22h → 1 sub-task    22h  (28th 3h 48m · 30th 8h · 31st 8h · 3rd 2h 12m)
Planned Leaves            leave and holidays, one sub-task each — never packed
```

- **Days are never split.** Walking a ticket's days in date order, when the next
  whole day would take the sub-task past 24h it closes at its real total and the
  next one starts. So a sub-task is `17h` or `23h 54m` — exactly the days it
  holds.
- **24h is a soft cap.** A single day worth more than that cannot be divided, so
  it stands alone.
- **One work log per day.** A 3-day sub-task gets three work logs, each dated
  and sized to its own day, so the Worklog Tracker still reads day by day. Only
  the sub-task count changes, never the daily figures.
- **Leave and holidays are never packed** — grouping them would lose which day
  was which kind of leave.
- **Splitting a day across tickets happens first**, so a day shared 60/40 gives
  its 60% to one ticket's group and its 40% to the other's.

### Naming

A sub-task is named for its **Task** plus the **title**, and nothing else:

| | |
| --- | --- |
| no title | `Development` |
| titled | `Development-Invoice totals` |

No dates in the name — the work logs say which days far better than a name can.

You type the title in step 4, while you remember what the work was, and it
travels with those days. Because a sub-task can hold days you titled
differently, step 5 does not guess: it lists the titles its own days carry with
the hours each covers, takes the biggest as the default, and lets you pick
another, type a new one (**Type another title…**), or drop it (**No title —
just the Task**). Where a ticket packed into several sub-tasks, **use the first
title for all** copies one name across them. Titles survive rebuilding the
plan, and the work log comment follows the title you settled on.

Leave sub-tasks can be named too (`Casual Leave-Eid holiday`) — the Task field
still carries the leave type, only the name changes.

### What each sub-task gets

- **Assignee**: you
- **Task** field: the Task you picked (the leave type on leave days)
- **Original Estimate**: everything it holds
- **Work logs**: one per day, dated and sized to that day
- **Status**: walked **To Do → In Progress → Done**, ticking the checklist that
  matches the Task (Development → Development Checklist)

### Why a run is quick

Almost all of a run is waiting for Jira, so the waiting is overlapped:

| | |
| --- | --- |
| **Sub-tasks** | up to `RUN_WORKERS` at once — they are independent |
| **The days inside one** | up to `LOG_WORKERS` at once — separate work logs against one issue. Each answers for itself, so a day that lands is remembered even if another fails |
| **Connections** | one pool of `POOL_SIZE`, so a parallel run reuses sockets instead of paying for a TLS handshake per call |
| **Reads** | retried twice on a 429 or a 5xx. Creates and work logs are **never** retried — the first attempt may have landed, and a duplicate work log is worse than an error |
| **The workflow** | the first sub-task of a project reads how to reach *Done*; the rest reuse it. If a reused transition is refused, the slow certain path runs instead, so it can only cost time, never correctness |
| **The sprint lookup** | teams spell sprints differently, so several spellings are asked at once rather than one after another |
| **The Task list** | fetched once per project, warmed in the background as soon as the sprint loads |
| **The attendance token** | held while valid, so only the first fetch needs a browser |

Measured against a Jira answering in 0.3s, a release of 20 days across two
tickets: **5.1s → 2.1s**. A plan of 18 separate sub-tasks: **5.7s → 4.5s**,
with 20 fewer calls. The sprint lookup: **1.8s → 0.3s**.

---

## 6. The command line

```powershell
python jira_logging_utility.py --dry-run     # preview, create nothing
python jira_logging_utility.py               # the same run, for real
```

Interactively it asks, in order: username, password/PAT, sprint, which issue,
start and end dates, how hours are set, the attendance sheet, which Task, and a
title. Then it prints the plan, asks you to confirm, and offers to log another
task afterwards.

The attendance sheet here is a file, not the portal: export it from the portal
as `.xlsx` for the range you want. When asked you can drag the file onto the
terminal, press Enter to open a file-browse dialog, or paste the path. The
**Date** and **Total Hours** columns are found automatically.

| Flag | What it does |
| --- | --- |
| `--dry-run` | Print the plan; create nothing |
| `--base-url URL` | Jira base URL |
| `--username NAME` | Username; leave it out and a blank username means "the password is a PAT" |
| `--sprint ST12-26.8` | Pre-fill the sprint |
| `--parent ST12-5876` | Skip sprint selection and use this issue |
| `--board 308` | Pin the sprint lookup to a board (normally detected) |
| `--start 2026-07-24 --end 2026-08-01` | Pre-fill the range |
| `--attendance-file PATH` | Use this `.xlsx`/`.csv` and skip the prompt |
| `--static-hours` | Flat 8h on full days; with no sheet, a plain 8h per weekday |
| `--task-value Development` | Skip the Task menu |
| `--summary "Edit Invoice Screen"` | Skip the title prompt |
| `--no-pack` | One sub-task per day, ungrouped |
| `--pack-cap 8` | Hours one grouped sub-task may hold (default 24) |
| `--ignore-logged` | Don't read your existing worklogs; log the full day regardless |
| `--leave-parent ST12-5524` | Force the Planned Leaves parent |
| `--leave-hours 8h` | Hours a leave/holiday sub-task gets |
| `--status-path "In Progress,Done"` | The statuses to walk |
| `--original-estimate 9h` / `--hours 8h` | Estimate and time when no sheet is used |
| `--no-assign-self` | Don't assign the sub-tasks to you |
| `--no-check-all` | Don't tick the checklist |
| `--include-weekends` | Also create Sat/Sun tasks in flat mode |
| `--no-verify-ssl` | Skip SSL verification (internal certificates) |

Fully non-interactive except the login:

```powershell
python jira_logging_utility.py --attendance-file attendance.xlsx `
  --parent ST12-5876 --start 2026-07-24 --end 2026-08-01 `
  --task-value Development --summary "Edit Invoice Screen" --dry-run
```

---

## 7. How it is put together

| File | Does |
| --- | --- |
| `jira_logging_utility.py` | The engine and the CLI: the Jira client, the hour rules, the timesheet read-back, grouping and packing |
| `app.py` | The web UI's backend — a thin Flask layer over that engine |
| `static/index.html` | The whole page: markup, styles and script in one file |
| `attendance_portal.py` | Fetches attendance from the portal's API |
| `sso_login.py` | Automatic sign-in: saved password → browser cookies → headless SSO |
| `jira_credentials.py` | Keeps your sign-in in Windows Credential Manager |

The rules live in exactly one place. `classify_day()` decides what a day is
worth, `remaining_day()` subtracts what Jira already has, and `pack_plan()`
groups the result — the CLI and the web UI both go through them, so the two can
never drift apart.

The backend endpoints, if you want to script against it: `/api/login`,
`/api/sso/{start,status,cancel}`, `/api/sprint`, `/api/tasks`,
`/api/attendance{,/fetch,/status,/cancel}`, `/api/logged`, `/api/days`,
`/api/plan`, `/api/run`, `/api/session`, `/api/reset`.

---

## 8. When something goes wrong

- **"Wrong username or password (HTTP 401)"** — the credentials are wrong, or
  password login is disabled (common on Jira Server). Create a **Personal
  Access Token**, leave the username blank, and paste the token as the password.
- **"Jira is asking for a CAPTCHA"** — too many failed logins locked the form.
  Sign in once in a browser to clear it, or use a PAT.
- **"Nothing saved yet for …"** on the Microsoft button — sign in once with your
  password so it can be saved; after that the button needs no typing.
- **"Sprint not found"** — usually a misspelt name. If your Jira has the sprint
  picker disabled, pass the board id (`--board <id>`, from the board URL
  `…RapidBoard.jspa?rapidView=<id>`).
- **The attendance fetch says the portal isn't signed in** — open
  `https://attendance.i2cinc.com/employee/attendance`, sign in with Microsoft
  once, then fetch again; or use **Open the portal in a window**. After that it
  stays signed in and runs in the background.
- **"Couldn't read your Jira worklogs"** — the check needs the `worklogAuthor` /
  `worklogDate` JQL fields, which a few instances restrict. Nothing is
  subtracted when it fails, so check the Worklog Tracker yourself — or pass
  `--ignore-logged` on the CLI.
- **A day you did log is still offered** — worklogs are matched to *you* and to
  the day they are `started` on. Time logged by someone else on your behalf, or
  dated to another day, is neither. Press **Re-check Jira** after logging
  anything in Jira directly.
- **"Couldn't find 'Date' and 'Total Hours' columns"** — the export headers
  differ from expected; open the file and check the header row.
- **"SSL certificate check failed"** — internal certificate: `--no-verify-ssl`.
- **The Done step fails on a required field** — that screen wants a field the
  tool cannot fill; note the field name so it can be added.
- **A leave day is not picked up as leave** — the portal labels it differently;
  note the exact text so it can be added.
- **`openpyxl` warns about a default style** — harmless.
