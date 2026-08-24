#!/usr/bin/env python3
"""
jira_credentials.py — remember the Jira sign-in so the one-click button works.

Sign in once with your username and password (or a PAT) and it is kept here;
after that the app signs itself in with no typing and no browser.

Where it is kept, best first:

  1. **Windows Credential Manager** via `keyring` — the password is held by the
     OS under your Windows account, the same place Edge keeps its passwords.
  2. A local file next to this script, if `keyring` isn't installed. That file
     is only base64 — **obfuscated, not encrypted**. Anyone who can read your
     files can read it, so keep `keyring` installed if you can.

Nothing is ever sent anywhere except to your own Jira.
"""

from __future__ import annotations

import base64
import json
import os
import stat
from typing import Optional
from urllib.parse import urlparse

SERVICE = "jira-timesheet"
HERE = os.path.dirname(os.path.abspath(__file__))
FALLBACK_FILE = os.path.join(HERE, ".jira-credentials")


def _key(base_url: str) -> str:
    """One saved sign-in per Jira host."""
    return (urlparse(base_url or "").hostname or base_url or "jira").lower()


def _keyring():
    try:
        import keyring  # noqa: WPS433
        return keyring
    except Exception:  # noqa: BLE001 — no keyring, or no usable backend
        return None


# --------------------------------------------------------------------------- #
#  File fallback (obfuscated only — see the module docstring)
# --------------------------------------------------------------------------- #
def _read_file() -> dict:
    try:
        with open(FALLBACK_FILE, "rb") as fh:
            return json.loads(base64.b64decode(fh.read()).decode("utf-8"))
    except Exception:  # noqa: BLE001 — missing or unreadable is just "nothing saved"
        return {}


def _write_file(store: dict) -> None:
    blob = base64.b64encode(json.dumps(store).encode("utf-8"))
    with open(FALLBACK_FILE, "wb") as fh:
        fh.write(blob)
    try:
        os.chmod(FALLBACK_FILE, stat.S_IRUSR | stat.S_IWUSR)   # owner only
    except OSError:
        pass


# --------------------------------------------------------------------------- #
#  API
# --------------------------------------------------------------------------- #
def save(base_url: str, username: str, password: str) -> str:
    """Remember this sign-in. Returns where it went, for the UI to report."""
    if not password:
        return ""
    payload = json.dumps({"username": username or "", "password": password})
    kr = _keyring()
    if kr is not None:
        try:
            kr.set_password(SERVICE, _key(base_url), payload)
            return "Windows Credential Manager"
        except Exception:  # noqa: BLE001 — fall through to the file
            pass
    store = _read_file()
    store[_key(base_url)] = payload
    _write_file(store)
    return "a local file"


def load(base_url: str) -> Optional[dict]:
    """The saved {'username', 'password'} for this Jira, or None."""
    kr = _keyring()
    if kr is not None:
        try:
            raw = kr.get_password(SERVICE, _key(base_url))
            if raw:
                return json.loads(raw)
        except Exception:  # noqa: BLE001
            pass
    raw = _read_file().get(_key(base_url))
    if raw:
        try:
            return json.loads(raw)
        except ValueError:
            return None
    return None

