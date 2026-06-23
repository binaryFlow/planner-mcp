"""
Browser-based Microsoft Graph auth for locked-down tenants.

This tenant blocks user consent and blocks generic first-party apps from Graph
(AADSTS65002 / admin-consent). The ONLY client authorized for our Tasks scopes
is Graph Explorer (de8bc8b5...), a single-page app that only mints tokens inside
a browser via MSAL.js.

So we serve a tiny stub page AT the Graph Explorer URL (the registered redirect
URI / SPA origin) inside a real Chromium, load the vendored MSAL.js into it, and
let MSAL do the work — exactly as Graph Explorer does:

    python auth_browser.py login    # one-time: opens an MSAL popup to sign in (+MFA)
    python auth_browser.py token    # headless: MSAL ssoSilent -> prints a fresh token

The signed-in session is persisted in .browser_profile/, so `token` runs with no
UI afterwards until the Entra session expires, at which point re-run `login`.
"""

import json
import os
import sys

from playwright.sync_api import sync_playwright

try:
    from dotenv import load_dotenv
except Exception:
    def load_dotenv(*_a, **_k):
        return False

_HERE = os.path.dirname(os.path.abspath(__file__))
load_dotenv(os.path.join(_HERE, ".env"))

PROFILE_DIR = os.path.join(_HERE, ".browser_profile")
MSAL_JS_PATH = os.path.join(_HERE, "vendor", "msal-browser.min.js")

CLIENT_ID = os.getenv("CLIENT_ID", "de8bc8b5-d9f9-48b1-a8ad-b748da725064")
TENANT_ID = os.getenv("TENANT_ID", "organizations")
LOGIN_HINT = os.getenv("GRAPH_LOGIN_HINT", "")
AUTHORITY = f"https://login.microsoftonline.com/{TENANT_ID}"

# Graph Explorer's redirect URI is window.location.href.toLowerCase(); this is
# its production URL and a registered SPA redirect. We serve our own stub page
# here so MSAL has a valid, same-origin context to run in.
REDIRECT_URI = os.getenv(
    "GRAPH_REDIRECT_URI", "https://developer.microsoft.com/en-us/graph/graph-explorer"
)

SCOPES = os.getenv("GRAPH_SCOPES", "User.Read Tasks.ReadWrite").split()

STUB_HTML = "<!doctype html><html><head><meta charset='utf-8'><title>auth</title></head><body>ok</body></html>"

# Runs in the page (window.msal = vendored MSAL.js). Returns the token or an error.
_JS = """
async ({clientId, authority, redirectUri, scopes, loginHint, interactive}) => {
  try {
    const pca = new msal.PublicClientApplication({
      auth: { clientId, authority, redirectUri, navigateToLoginRequestUrl: false },
      cache: { cacheLocation: 'localStorage' },
    });
    const req = { scopes };
    if (loginHint) req.loginHint = loginHint;
    let res;
    if (interactive) {
      res = await pca.acquireTokenPopup({ ...req, prompt: 'select_account' });
    } else {
      res = await pca.ssoSilent(req);
    }
    return {
      ok: true,
      accessToken: res.accessToken,
      expiresOn: res.expiresOn ? res.expiresOn.getTime() : null,
    };
  } catch (e) {
    return {
      ok: false,
      error: (e && e.errorCode) || 'error',
      error_description: (e && (e.errorMessage || e.message)) || String(e),
    };
  }
}
"""


def _run_flow(interactive: bool) -> dict:
    if not os.path.exists(MSAL_JS_PATH):
        raise RuntimeError(
            "Missing vendor/msal-browser.min.js. Fetch it once:\n"
            "  mkdir -p vendor && curl -fsSL -o vendor/msal-browser.min.js "
            "https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js"
        )
    with open(MSAL_JS_PATH, "r") as fh:
        msal_js = fh.read()

    with sync_playwright() as p:
        context = p.chromium.launch_persistent_context(
            PROFILE_DIR,
            headless=not interactive,
            args=["--no-first-run", "--no-default-browser-check"],
        )
        # No hard timeout for interactive login (user may take a while with MFA).
        context.set_default_timeout(0 if interactive else 60000)

        def handle(route):
            url = route.request.url
            if url.lower().startswith(REDIRECT_URI):
                route.fulfill(status=200, content_type="text/html", body=STUB_HTML)
            else:
                route.continue_()

        # Context-level so popups and iframes (the OAuth redirect target) are
        # stubbed too — keeps Graph Explorer's own app from loading and racing
        # us to redeem the single-use auth code.
        context.route("https://developer.microsoft.com/**", handle)

        page = context.pages[0] if context.pages else context.new_page()
        page.goto(REDIRECT_URI, wait_until="domcontentloaded")
        page.add_script_tag(content=msal_js)

        result = page.evaluate(
            _JS,
            {
                "clientId": CLIENT_ID,
                "authority": AUTHORITY,
                "redirectUri": REDIRECT_URI,
                "scopes": SCOPES,
                "loginHint": LOGIN_HINT,
                "interactive": interactive,
            },
        )
        context.close()
        return result


def main() -> int:
    mode = sys.argv[1] if len(sys.argv) > 1 else "token"
    interactive = mode == "login"

    try:
        result = _run_flow(interactive=interactive)
    except Exception as exc:
        result = {"ok": False, "error": "exception", "error_description": str(exc)}

    if not result.get("ok"):
        msg = result.get("error_description") or result.get("error") or "unknown error"
        if interactive:
            print(f"Login failed: {msg}", file=sys.stderr)
        else:
            print(json.dumps({"error": result.get("error", "error"), "error_description": msg}))
        return 1

    if interactive:
        print("\n✅ Signed in. The MCP server will now fetch tokens automatically.")
    else:
        expires_on = result.get("expiresOn")
        import time
        expires_in = int(expires_on / 1000 - time.time()) if expires_on else 3600
        print(json.dumps({"access_token": result["accessToken"], "expires_in": max(expires_in, 60)}))
    return 0


if __name__ == "__main__":
    sys.exit(main())
