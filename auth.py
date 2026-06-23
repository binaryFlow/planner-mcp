"""
Token provider for the Planner MCP server.

Dispatches to one of two auth backends based on AUTH_METHOD in .env:

  * "device" (default) -> auth_device.py: MSAL device-code flow with your own
    Azure app registration. Supported, lightweight, self-refreshing.

  * "browser" -> auth_browser.py: drives a headless Chromium that replays Graph
    Explorer's browser sign-in. The no-admin fallback for locked-down tenants
    that block app consent. Heavier (needs Playwright) and unsupported. The
    browser fetch runs in a SUBPROCESS so Playwright's sync API never collides
    with the MCP server's asyncio event loop; tokens are cached on disk so the
    browser only re-launches near expiry.
"""

import json
import os
import subprocess
import sys
import time

try:
    from dotenv import load_dotenv
except Exception:
    def load_dotenv(*_a, **_k):
        return False

_HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_HERE, ".env"))

AUTH_METHOD = os.getenv("AUTH_METHOD", "device").lower()
CACHE_FILE = os.path.join(_HERE, ".token_cache.json")
BROWSER_HELPER = os.path.join(_HERE, "auth_browser.py")
_SKEW = 120  # refresh when fewer than this many seconds remain


def _read_cache() -> dict:
    try:
        with open(CACHE_FILE, "r") as fh:
            return json.load(fh)
    except Exception:
        return {}


def _write_cache(token: str, expires_at: float) -> None:
    with open(CACHE_FILE, "w") as fh:
        json.dump({"access_token": token, "expires_at": expires_at}, fh)
    try:
        os.chmod(CACHE_FILE, 0o600)
    except OSError:
        pass


def _fetch_via_browser(timeout: int = 90) -> str | None:
    """Run the headless Playwright flow in a subprocess; return a fresh token."""
    try:
        proc = subprocess.run(
            [sys.executable, BROWSER_HELPER, "token"],
            capture_output=True,
            text=True,
            timeout=timeout,
            cwd=_HERE,
        )
    except subprocess.TimeoutExpired:
        return None

    out = (proc.stdout or "").strip()
    if not out:
        return None
    try:
        data = json.loads(out.splitlines()[-1])
    except Exception:
        return None
    token = data.get("access_token")
    if not token:
        return None
    expires_in = int(data.get("expires_in", 3600))
    _write_cache(token, time.time() + expires_in)
    return token


def _browser_token() -> str | None:
    cache = _read_cache()
    token = cache.get("access_token")
    if token and cache.get("expires_at", 0) - _SKEW > time.time():
        return token
    return _fetch_via_browser()


def acquire_token_silent() -> str | None:
    """Return a valid access token using the configured auth backend."""
    if AUTH_METHOD == "browser":
        return _browser_token()
    from auth_device import get_token

    return get_token()
