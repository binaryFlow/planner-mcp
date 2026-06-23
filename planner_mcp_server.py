"""
Microsoft Planner MCP Server

Exposes Planner tasks as tools callable from GitHub Copilot / VS Code.

Requirements:
    pip install -r requirements.txt

Setup (one-time):
    1. Copy .env.example to .env and configure auth (see README.md).
    2. Run:  python login.py   and sign in once.

After that the server obtains and refreshes Microsoft Graph access tokens
automatically (see auth.py). Re-run login.py only if you ever see "Not signed in".
"""

import json
import os
import re

import requests
from dotenv import load_dotenv
from mcp.server.fastmcp import FastMCP

load_dotenv(os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env"))

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
TENANT_ID = os.getenv("TENANT_ID", "")

mcp = FastMCP("Microsoft Planner")


def _get_token() -> str:
    # Preferred: silently mint a fresh token from the MSAL cache (see auth.py).
    # Sign in once with `python login.py`; refreshes happen automatically after.
    try:
        from auth import acquire_token_silent

        token = acquire_token_silent()
        if token:
            return token
    except Exception:
        pass

    # Fallback: a manually pasted Graph Explorer token in .env (legacy path).
    token = os.getenv("ACCESS_TOKEN", "")
    if token:
        return token

    raise RuntimeError(
        "Not signed in. Run a one-time login:  python login.py\n"
        "(After that the server refreshes tokens automatically.)"
    )


def _graph_get(token: str, path: str) -> dict:
    resp = requests.get(
        f"{GRAPH_BASE}{path}",
        headers={"Authorization": f"Bearer {token}"},
        timeout=15,
    )
    resp.raise_for_status()
    return resp.json()


def _bucket_name(token: str, bucket_id: str, cache: dict) -> str:
    if bucket_id not in cache:
        try:
            cache[bucket_id] = _graph_get(token, f"/planner/buckets/{bucket_id}")["name"]
        except Exception:
            cache[bucket_id] = bucket_id
    return cache[bucket_id]


def _plan_categories(token: str, plan_id: str, cache: dict) -> dict:
    """Return the {category key: label} map for a plan (category1..category25)."""
    if plan_id not in cache:
        try:
            details = _graph_get(token, f"/planner/plans/{plan_id}/details")
            cache[plan_id] = details.get("categoryDescriptions", {}) or {}
        except Exception:
            cache[plan_id] = {}
    return cache[plan_id]


def _task_labels(token: str, task: dict, cache: dict) -> list:
    """Resolve a task's appliedCategories to their human-readable label names."""
    applied = task.get("appliedCategories", {}) or {}
    active = [key for key, on in applied.items() if on]
    if not active:
        return []
    categories = _plan_categories(token, task.get("planId", ""), cache)
    return [categories.get(key) or key for key in active]


def _plan_group_id(token: str, plan_id: str, cache: dict) -> str:
    """Return the owning group (Microsoft 365 group) id for a plan."""
    if plan_id not in cache:
        try:
            plan = _graph_get(token, f"/planner/plans/{plan_id}")
            container = plan.get("container", {}) or {}
            cache[plan_id] = container.get("containerId") or plan.get("owner", "")
        except Exception:
            cache[plan_id] = ""
    return cache[plan_id]


def _checklist(details: dict) -> list:
    """Flatten a task details checklist dict into an ordered list of items."""
    items = (details.get("checklist", {}) or {}).items()
    ordered = sorted(items, key=lambda kv: (kv[1] or {}).get("orderHint", ""))
    return [
        {"title": v.get("title", ""), "isChecked": bool(v.get("isChecked"))}
        for _, v in ordered
        if v
    ]


def _strip_html(html: str) -> str:
    """Reduce an HTML email/post body to readable plain text."""
    text = re.sub(r"<\s*br\s*/?>", "\n", html, flags=re.IGNORECASE)
    text = re.sub(r"</\s*p\s*>", "\n", text, flags=re.IGNORECASE)
    text = re.sub(r"<[^>]+>", "", text)
    text = re.sub(r"&nbsp;", " ", text)
    text = re.sub(r"&amp;", "&", text)
    text = re.sub(r"&lt;", "<", text)
    text = re.sub(r"&gt;", ">", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def _task_comments(token: str, task: dict, group_cache: dict) -> list:
    """Fetch the task's conversation thread posts (comments) as plain text."""
    thread_id = task.get("conversationThreadId")
    if not thread_id:
        return []
    group_id = _plan_group_id(token, task.get("planId", ""), group_cache)
    if not group_id:
        return []
    try:
        posts = _graph_get(
            token, f"/groups/{group_id}/threads/{thread_id}/posts"
        ).get("value", [])
    except Exception:
        return []
    comments = []
    for post in posts:
        sender = ((post.get("from", {}) or {}).get("emailAddress", {}) or {})
        comments.append({
            "from": sender.get("name") or sender.get("address") or "",
            "created": (post.get("createdDateTime", "") or "")[:19],
            "content": _strip_html((post.get("body", {}) or {}).get("content", "")),
        })
    return comments


@mcp.tool()
def get_planner_tasks(bucket: str = "", priority: str = "", with_details: bool = False) -> str:
    """
    Fetch all tasks assigned to me, grouped by bucket.
    Optionally filter by bucket name (case-insensitive partial match).
    Optionally filter by priority: urgent, important, medium, low.
    Set with_details=true to include each task's description, checklist, and comment/conversation
    thread (slower — extra API calls per task).
    Returns title, priority, progress (not_started / in_progress / completed), due date, assignee
    count, and labels per task; plus description, checklist, and comments when with_details=true.
    """
    token = _get_token()
    tasks = _graph_get(token, "/me/planner/tasks").get("value", [])

    def _progress(pct: int) -> str:
        if pct == 100:
            return "completed"
        if pct == 50:
            return "in_progress"
        return "not_started"

    _priority_map = {0: "urgent", 1: "urgent", 2: "important", 3: "important",
                     4: "medium", 5: "medium", 6: "medium", 7: "low", 8: "low", 9: "low"}

    def _priority_label(p: int) -> str:
        return _priority_map.get(p, "medium")

    bucket_cache: dict = {}
    category_cache: dict = {}
    group_cache: dict = {}
    by_bucket: dict[str, list] = {}
    for task in tasks:
        name = _bucket_name(token, task.get("bucketId", ""), bucket_cache)
        if bucket and bucket.lower() not in name.lower():
            continue
        task_priority = _priority_label(task.get("priority", 5))
        if priority and priority.lower() != task_priority:
            continue
        task_id = task["id"]
        entry: dict = {
            "id": task_id,
            "title": task["title"],
            "priority": task_priority,
            "progress": _progress(task.get("percentComplete", 0)),
            "due": task.get("dueDateTime", "")[:10] if task.get("dueDateTime") else None,
            "assignees": len(task.get("assignments", {})),
            "labels": _task_labels(token, task, category_cache),
            "url": f"https://tasks.office.com/{TENANT_ID}/Home/Task/{task_id}" if TENANT_ID else None,
        }
        if with_details:
            try:
                details = _graph_get(token, f"/planner/tasks/{task['id']}/details")
                description = details.get("description", "").strip()
                if description:
                    entry["description"] = description
                checklist = _checklist(details)
                if checklist:
                    entry["checklist"] = checklist
            except Exception:
                pass
            comments = _task_comments(token, task, group_cache)
            if comments:
                entry["comments"] = comments
        by_bucket.setdefault(name, []).append(entry)

    return json.dumps(by_bucket, indent=2, ensure_ascii=False)


@mcp.tool()
def update_planner_tasks(task_ids: list[str], status: str) -> str:
    """
    Update the progress of one or more Planner tasks.
    task_ids: list of Planner task IDs (the 'id' field from get_planner_tasks).
    status: one of 'not_started', 'in_progress', or 'completed'.
    Returns a summary of updated and failed tasks.
    """
    status_map = {"not_started": 0, "in_progress": 50, "completed": 100}
    if status not in status_map:
        return json.dumps({"error": f"Invalid status '{status}'. Use: not_started, in_progress, completed"})

    percent = status_map[status]
    token = _get_token()
    results: dict[str, str] = {}

    for task_id in task_ids:
        try:
            # Fetch the current task to get the required ETag for optimistic concurrency.
            task_resp = requests.get(
                f"{GRAPH_BASE}/planner/tasks/{task_id}",
                headers={"Authorization": f"Bearer {token}"},
                timeout=15,
            )
            task_resp.raise_for_status()
            etag = task_resp.headers.get("ETag") or task_resp.json().get("@odata.etag", "")

            patch_resp = requests.patch(
                f"{GRAPH_BASE}/planner/tasks/{task_id}",
                headers={
                    "Authorization": f"Bearer {token}",
                    "Content-Type": "application/json",
                    "If-Match": etag,
                },
                json={"percentComplete": percent},
                timeout=15,
            )
            patch_resp.raise_for_status()
            results[task_id] = "updated"
        except requests.HTTPError as exc:
            results[task_id] = f"error: {exc.response.status_code} {exc.response.text[:200]}"
        except Exception as exc:
            results[task_id] = f"error: {exc}"

    return json.dumps(results, indent=2, ensure_ascii=False)


@mcp.tool()
def get_planner_buckets() -> str:
    """List all bucket names from tasks assigned to me."""
    token = _get_token()
    tasks = _graph_get(token, "/me/planner/tasks").get("value", [])

    bucket_cache: dict = {}
    bucket_ids = {t["bucketId"] for t in tasks if t.get("bucketId")}
    names = [_bucket_name(token, bid, bucket_cache) for bid in bucket_ids]

    return json.dumps(sorted(names), ensure_ascii=False)


if __name__ == "__main__":
    mcp.run()
