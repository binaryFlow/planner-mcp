"""
One-time sign-in for the Planner MCP server.

Run once:
    python login.py

Behaviour depends on AUTH_METHOD in .env:
  * "device" (default) -> prints a URL + code; open it, sign in. MSAL then
    refreshes tokens automatically.
  * "browser"          -> opens a Chromium window with an MSAL sign-in popup
    (the no-admin fallback). Sign in + MFA; the session is saved locally.

Re-run only if the server later reports "Not signed in".
"""

import os
import subprocess
import sys

try:
    from dotenv import load_dotenv
except Exception:
    def load_dotenv(*_a, **_k):
        return False

_HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_HERE, ".env"))


def main() -> int:
    method = os.getenv("AUTH_METHOD", "device").lower()
    if method == "browser":
        return subprocess.call(
            [sys.executable, os.path.join(_HERE, "auth_browser.py"), "login"]
        )

    from auth_device import device_login

    result = device_login()
    name = (result.get("id_token_claims") or {}).get("preferred_username", "")
    print(f"\n✅ Signed in{f' as {name}' if name else ''}.")
    print("The MCP server will now obtain and refresh tokens automatically.")
    return 0


if __name__ == "__main__":
    sys.exit(main())
