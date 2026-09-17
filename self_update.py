#!/usr/bin/env python3
"""
self_update.py - Pull the latest code before the app starts.

Every run begins with a fetch: if the checkout is behind its remote, it is
fast-forwarded and the process re-executes itself so the new code is the code
that runs. Python has already read app.py by the time a pull lands, so without
that restart an update would sit on disk unused until the next launch.

Nothing here is allowed to stop the app. No git, no network, no remote, a
branch with no upstream, edits in the working tree - each one prints a line and
hands control straight back. The update is a convenience; the app is the point.

    import self_update
    self_update.update_and_restart()   # may not return - it re-execs

Two ways to switch it off: --no-update on the command line, or set
JIRA_LOGGER_NO_UPDATE=1 in the environment.
"""

from __future__ import annotations

import os
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))

# Set on the child we re-exec, so a repository that somehow reports "behind"
# twice in a row cannot spin the process forever. One restart per launch.
_RESTART_GUARD = "JIRA_LOGGER_UPDATED"

# Opt-out, for an offline run or while you are editing the code.
_DISABLE = "JIRA_LOGGER_NO_UPDATE"
_DISABLE_FLAG = "--no-update"

# A fetch over VPN is usually under a second; these are the point at which we
# stop waiting and just start the app.
FETCH_TIMEOUT = 20         # seconds - talks to the network
LOCAL_TIMEOUT = 10         # seconds - reads the local repository


def _git(*args, timeout=LOCAL_TIMEOUT):
    """
    Run one git command in this folder.

    Returns (ok, output). Never raises: a missing git, a timeout and a
    non-zero exit all come back the same way, because the caller treats them
    the same way - skip the update, start the app.
    """
    try:
        done = subprocess.run(
            ["git", *args],
            cwd=HERE,
            capture_output=True,
            text=True,
            timeout=timeout,
            # Windows: keep a console window from flashing up per git call.
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except (OSError, subprocess.SubprocessError) as exc:
        return False, str(exc)
    output = (done.stdout or "").strip() or (done.stderr or "").strip()
    return done.returncode == 0, output


def _opted_out(argv):
    """True when this launch should not touch git at all."""
    if os.environ.get(_RESTART_GUARD) == "1":
        # We are the restarted process; the pull already happened.
        return True
    if os.environ.get(_DISABLE) == "1":
        print(f"Update:  skipped - {_DISABLE}=1.")
        return True
    if _DISABLE_FLAG in argv:
        print(f"Update:  skipped - {_DISABLE_FLAG}.")
        return True
    return False


def _restart():
    """Replace this process with a fresh one running the updated code."""
    env = dict(os.environ)
    env[_RESTART_GUARD] = "1"
    argv = [sys.executable, os.path.abspath(sys.argv[0]), *sys.argv[1:]]
    print("Update:  restarting with the new code ...\n", flush=True)
    # execv keeps the same console and the same exit code on POSIX. On Windows
    # it spawns and the parent dies first, which confuses a shell that is
    # waiting on us, so there we wait on a child instead.
    if os.name == "nt":
        raise SystemExit(subprocess.call(argv, cwd=HERE, env=env))
    os.execve(sys.executable, argv, env)


def update_and_restart(argv=None):
    """
    Fast-forward the checkout, then re-exec if anything came down.

    Returns True when the code is current (pulled or already up to date) and
    False when the update was skipped for any reason. On a successful pull it
    does not return at all - the process is replaced.
    """
    argv = sys.argv[1:] if argv is None else list(argv)

    if _opted_out(argv):
        return False

    ok, _ = _git("rev-parse", "--is-inside-work-tree")
    if not ok:
        print("Update:  skipped - not a git checkout (or git is not installed).")
        return False

    ok, upstream = _git("rev-parse", "--abbrev-ref", "--symbolic-full-name", "@{u}")
    if not ok:
        print("Update:  skipped - this branch does not track a remote.")
        return False

    # Edits to tracked files are yours. A pull could rewrite them or stop
    # halfway, so we leave the tree exactly as we found it and say so.
    # Untracked files are none of a fast-forward's business - counting them
    # would let one stray __pycache__ switch updates off for good.
    ok, dirty = _git("status", "--porcelain", "--untracked-files=no")
    if ok and dirty:
        print("Update:  skipped - you have uncommitted changes to tracked files.")
        return False

    print(f"Update:  checking {upstream} for new code ...", flush=True)
    ok, message = _git("fetch", "--quiet", timeout=FETCH_TIMEOUT)
    if not ok:
        print(f"Update:  skipped - could not reach the remote ({message or 'no detail'}).")
        return False

    ok, local = _git("rev-parse", "HEAD")
    fine, remote = _git("rev-parse", "@{u}")
    if ok and fine and local == remote:
        print("Update:  already up to date.")
        return True

    # --ff-only: if the branch has drifted from the remote, stop rather than
    # start a merge nobody asked for and leave the tree mid-conflict.
    ok, message = _git("merge", "--ff-only", "@{u}", timeout=FETCH_TIMEOUT)
    if not ok:
        print(f"Update:  skipped - could not fast-forward ({message or 'no detail'}).")
        print("         Run 'git pull' yourself to see what is in the way.")
        return False

    _, head = _git("log", "-1", "--pretty=%h %s")
    print(f"Update:  pulled the latest changes -> {head}")
    _restart()
    return True


if __name__ == "__main__":
    update_and_restart()
