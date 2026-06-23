# Microsoft Planner MCP Server

An [MCP](https://modelcontextprotocol.io) server that exposes your **Microsoft
Planner** tasks to MCP clients (Claude, GitHub Copilot, VS Code, and others).
List your tasks grouped by bucket, drill into details/checklists/comments, and
update task progress, all from your assistant.

## Tools

| Tool | What it does |
| --- | --- |
| `get_planner_tasks` | Tasks assigned to you, grouped by bucket. Filter by bucket/priority; `with_details=true` adds description, checklist, and comments. |
| `update_planner_tasks` | Set progress (`not_started` / `in_progress` / `completed`) for one or more tasks. |
| `get_planner_buckets` | List your bucket names. |

## Install

```bash
python -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cp .env.example .env        # then edit .env
```

## Authentication

This server talks to Microsoft Graph with **delegated** permissions
(`User.Read`, `Tasks.ReadWrite`). Pick one of two methods via `AUTH_METHOD` in
`.env`.

### Method 1: device code (default, recommended)

Uses **your own** Azure app registration. One-time portal setup:

1. Microsoft Entra ID, then **App registrations**, then **New registration**.
2. **Authentication**: add a **Mobile and desktop applications** platform and
   enable **Allow public client flows**.
3. **API permissions**: add delegated Microsoft Graph permissions
   `User.Read` and `Tasks.ReadWrite` (and grant/consent for your account).
4. Put the IDs in `.env`:

   ```ini
   AUTH_METHOD=device
   CLIENT_ID=<application-client-id>
   TENANT_ID=<directory-tenant-id>
   ```

Then sign in once:

```bash
python login.py     # prints a URL + code; open it and sign in
```

MSAL caches a refresh token and the server refreshes access tokens silently.

### Method 2: browser fallback (no admin, locked-down tenants)

> ⚠️ **Read this first.** This method drives a headless Chromium that replays
> Graph Explorer's sign-in, **borrowing the Graph Explorer first-party client**
> to obtain tokens. Use it only if your tenant blocks app consent and you have
> no admin access and no way to register an app. It is **unsupported**, depends
> on undocumented Graph Explorer behavior (and may break without notice), and
> may be **against your organization's policies**, so check before using it.
> When you can, use Method 1 instead.

```bash
pip install playwright && playwright install chromium
# fetch the MSAL.js library used inside the browser:
mkdir -p vendor && curl -fsSL -o vendor/msal-browser.min.js \
  https://cdn.jsdelivr.net/npm/@azure/msal-browser@2.38.4/lib/msal-browser.min.js
```

`.env`:

```ini
AUTH_METHOD=browser
TENANT_ID=<directory-tenant-id>
GRAPH_LOGIN_HINT=you@example.com   # optional
```

Then:

```bash
python login.py     # opens a browser window; sign in + MFA once
```

The signed-in session is saved in `.browser_profile/`, and the server fetches
tokens silently afterwards. Re-run `login.py` if the session ever lapses.

## Connect an MCP client

Example (VS Code / Claude `mcp.json`-style config):

```jsonc
{
  "mcpServers": {
    "planner": {
      "command": "/absolute/path/to/.venv/bin/python",
      "args": ["/absolute/path/to/planner_mcp_server.py"]
    }
  }
}
```

## Configuration reference

| Variable | Default | Notes |
| --- | --- | --- |
| `AUTH_METHOD` | `device` | `device` or `browser` |
| `CLIENT_ID` | (none) | Your app's client ID (device method) |
| `TENANT_ID` | `organizations` | Your tenant GUID |
| `GRAPH_SCOPES` | `User.Read Tasks.ReadWrite` | Space-separated delegated scopes |
| `GRAPH_LOGIN_HINT` | (none) | Pre-fill the sign-in account |

## Security notes

Never commit `.env`, `.token_cache.json`, `.msal_cache.json`, or
`.browser_profile/`, since they hold credentials. They are in `.gitignore`.

## Disclaimer

This is an independent, community project. It is not affiliated with, endorsed
by, or sponsored by Microsoft. "Microsoft", "Microsoft Planner", and "Microsoft
Graph" are trademarks of Microsoft Corporation, used here only to describe
interoperability.

## License

MIT. See [LICENSE](LICENSE).
