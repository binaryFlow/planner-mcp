"""
Device-code authentication (DEFAULT, recommended).

Uses MSAL with YOUR OWN Azure app registration. This is the supported path and
needs no browser automation — you sign in once via a device code, and MSAL then
silently refreshes tokens from its cache.

Setup (one-time, in the Azure portal — needs an app registration you own):
    1. Register an app (Microsoft Entra ID -> App registrations -> New).
    2. Add a "Mobile and desktop" / public-client platform; enable
       "Allow public client flows".
    3. Add delegated Microsoft Graph permissions: User.Read, Tasks.ReadWrite.
    4. Copy the Application (client) ID and Directory (tenant) ID into .env:
           CLIENT_ID=<your-app-client-id>
           TENANT_ID=<your-tenant-id>

Then run `python login.py` once.

If you have NO admin access and your tenant blocks app consent, use the browser
fallback instead (set AUTH_METHOD=browser; see auth_browser.py).
"""

import os

import msal

try:
    from dotenv import load_dotenv
except Exception:
    def load_dotenv(*_a, **_k):
        return False

_HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_HERE, ".env"))

CACHE_FILE = os.path.join(_HERE, ".msal_cache.json")

CLIENT_ID = os.getenv("CLIENT_ID", "")
TENANT_ID = os.getenv("TENANT_ID", "organizations")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"
SCOPES = os.getenv("GRAPH_SCOPES", "User.Read Tasks.ReadWrite").split()


def _require_client_id() -> str:
    if not CLIENT_ID:
        raise RuntimeError(
            "CLIENT_ID is not set. Device-code auth needs your own Azure app "
            "registration — set CLIENT_ID (and TENANT_ID) in .env. See "
            "auth_device.py for setup, or use the no-admin browser fallback "
            "(AUTH_METHOD=browser)."
        )
    return CLIENT_ID


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as fh:
            cache.deserialize(fh.read())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as fh:
            fh.write(cache.serialize())
        try:
            os.chmod(CACHE_FILE, 0o600)
        except OSError:
            pass


def _app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        _require_client_id(), authority=AUTHORITY, token_cache=cache
    )


def get_token() -> str | None:
    """Return a fresh access token from the MSAL cache, or None if not signed in."""
    if not CLIENT_ID:
        return None
    cache = _load_cache()
    app = _app(cache)
    accounts = app.get_accounts()
    if not accounts:
        return None
    result = app.acquire_token_silent(SCOPES, account=accounts[0])
    _save_cache(cache)
    if result and "access_token" in result:
        return result["access_token"]
    return None


def device_login() -> dict:
    """Interactive one-time device-code sign-in."""
    cache = _load_cache()
    app = _app(cache)
    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise RuntimeError(f"Failed to start device flow: {flow}")
    print(flow["message"], flush=True)
    result = app.acquire_token_by_device_flow(flow)  # blocks until you finish
    _save_cache(cache)
    if "access_token" not in result:
        raise RuntimeError(
            f"Login failed: {result.get('error_description', result)}"
        )
    return result
