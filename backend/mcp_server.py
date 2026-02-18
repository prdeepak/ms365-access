"""MS365 Access MCP Server.

Exposes Microsoft 365 services (Mail, Calendar, OneDrive, SharePoint)
as MCP tools for use with Claude Code, Claude Desktop, Cursor, etc.

Thin wrapper around the ms365-access REST API (localhost:8365).

Usage:
    # stdio transport (for Claude Code / Claude Desktop)
    python -m mcp_server

    # streamable-http transport (for remote clients)
    python -m mcp_server --transport streamable-http --port 8366
"""

import argparse
import json
import logging
import os

import httpx
from mcp.server.fastmcp import FastMCP

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

_host = os.environ.get("MCP_HOST", "127.0.0.1")
mcp = FastMCP("ms365-access", host=_host, port=8366)

BASE_URL = os.environ.get("MS365_API_URL", "http://localhost:8365")
API_KEY = os.environ.get("MS365_API_KEY", "")


# ---------------------------------------------------------------------------
# HTTP helper
# ---------------------------------------------------------------------------

def _headers() -> dict:
    h = {"Content-Type": "application/json"}
    if API_KEY:
        h["Authorization"] = f"Bearer {API_KEY}"
    return h


def _get(path: str, params: dict | None = None) -> dict | list | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.get(path, params={k: v for k, v in (params or {}).items() if v is not None})
        r.raise_for_status()
        return r.json()


def _post(path: str, data: dict | None = None) -> dict | list | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.post(path, json=data)
        r.raise_for_status()
        return r.json()


def _patch(path: str, data: dict | None = None) -> dict | list | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.patch(path, json=data)
        r.raise_for_status()
        return r.json()


def _delete(path: str) -> dict | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.delete(path)
        r.raise_for_status()
        try:
            return r.json()
        except Exception:
            return {"deleted": True}


# ===========================================================================
# Mail Tools
# ===========================================================================

@mcp.tool()
def mail_list_folders() -> str:
    """List all mail folders (Inbox, Sent, Drafts, etc.).

    Returns folder names, IDs, and unread counts.
    """
    return json.dumps(_get("/mail/folders"), default=str)


@mcp.tool()
def mail_list_messages(
    folder: str | None = None,
    top: int = 25,
    search: str | None = None,
    filter: str | None = None,
) -> str:
    """List email messages, optionally filtered by folder or search.

    Args:
        folder: Folder ID or well-known name (inbox, archive, junkemail). Default: all messages.
        top: Max messages to return (default 25)
        search: Free-text search query
        filter: OData filter expression (e.g. "isRead eq false")
    """
    params = {"top": top, "search": search, "filter": filter}
    if folder:
        params["folder_id"] = folder
    return json.dumps(_get("/mail/messages", params), default=str)


@mcp.tool()
def mail_search(
    q: str,
    top: int = 25,
) -> str:
    """Search email messages by keyword.

    Args:
        q: Search query (searches subject, body, sender)
        top: Max results (default 25)
    """
    return json.dumps(_get("/mail/search", {"q": q, "top": top}), default=str)


@mcp.tool()
def mail_get_message(message_id: str) -> str:
    """Get a single email message by ID with full body.

    Args:
        message_id: Message ID
    """
    return json.dumps(_get(f"/mail/messages/{message_id}"), default=str)


@mcp.tool()
def mail_get_threads(
    folder: str | None = None,
    top: int = 25,
) -> str:
    """List email threads (conversations) grouped by conversationId.

    Args:
        folder: Folder ID or well-known name. Default: all.
        top: Max threads (default 25)
    """
    params = {"top": top}
    if folder:
        params["folder_id"] = folder
    return json.dumps(_get("/mail/threads", params), default=str)


@mcp.tool()
def mail_create_draft(
    subject: str,
    body: str = "",
    to_recipients: list[str] | None = None,
    cc_recipients: list[str] | None = None,
    body_type: str = "HTML",
) -> str:
    """Create a new draft email in the Drafts folder (does not send it).

    Requires write:draft permission. The returned message object contains the
    draft ID — use mail_update to edit it or (if permitted) mail_send_draft to send.

    Args:
        subject: Email subject
        body: Email body content (default empty)
        to_recipients: List of recipient email addresses (optional)
        cc_recipients: CC recipients (optional)
        body_type: 'HTML' or 'Text' (default 'HTML')
    """
    data: dict = {"subject": subject, "body": body, "body_type": body_type}
    if to_recipients:
        data["to_recipients"] = to_recipients
    if cc_recipients:
        data["cc_recipients"] = cc_recipients
    return json.dumps(_post("/mail/drafts", data), default=str)


@mcp.tool()
def mail_create_reply_draft(
    message_id: str,
    comment: str = "",
    reply_all: bool = False,
) -> str:
    """Create a draft reply to a message (does not send it).

    Requires write:draft permission. Returns the draft message object with its
    ID — use mail_update to edit the body before sending.

    Args:
        message_id: ID of the message to reply to
        comment: Reply text to prepend to the quoted thread (default empty)
        reply_all: Create reply-all draft instead of reply (default False)
    """
    url = f"/mail/messages/{message_id}/draftReply"
    if reply_all:
        url += "?reply_all=true"
    return json.dumps(_post(url, {"comment": comment} if comment else None), default=str)


@mcp.tool()
def mail_send(
    subject: str,
    body: str,
    to_recipients: list[str],
    cc_recipients: list[str] | None = None,
    body_type: str = "HTML",
) -> str:
    """Send an email message.

    Args:
        subject: Email subject
        body: Email body content
        to_recipients: List of recipient email addresses
        cc_recipients: CC recipients (optional)
        body_type: 'HTML' or 'Text' (default 'HTML')
    """
    data = {
        "subject": subject,
        "body": body,
        "to_recipients": to_recipients,
        "body_type": body_type,
    }
    if cc_recipients:
        data["cc_recipients"] = cc_recipients
    return json.dumps(_post("/mail/messages", data), default=str)


@mcp.tool()
def mail_reply(
    message_id: str,
    comment: str,
) -> str:
    """Reply to an email message.

    Args:
        message_id: ID of the message to reply to
        comment: Reply body text
    """
    return json.dumps(_post(f"/mail/messages/{message_id}/reply", {"comment": comment}), default=str)


@mcp.tool()
def mail_move(
    message_id: str,
    destination_folder: str,
) -> str:
    """Move a message to a different folder.

    Args:
        message_id: Message ID
        destination_folder: Destination folder ID or well-known name (archive, junkemail, deleteditems)
    """
    return json.dumps(_post(f"/mail/messages/{message_id}/move", {"destination_id": destination_folder}), default=str)


@mcp.tool()
def mail_update(
    message_id: str,
    is_read: bool | None = None,
    flag: str | None = None,
    body: str | None = None,
    body_type: str = "HTML",
) -> str:
    """Update message properties (mark read/unread, flag, body content).

    Args:
        message_id: Message ID
        is_read: Mark as read (True) or unread (False)
        flag: Flag status - 'flagged', 'complete', or 'notFlagged'
        body: New body content (HTML or plain text)
        body_type: 'HTML' or 'Text' (default 'HTML')
    """
    data = {}
    if is_read is not None:
        data["isRead"] = is_read
    if flag:
        data["flag"] = {"flagStatus": flag}
    if body is not None:
        data["body"] = body
        data["body_type"] = body_type
    return json.dumps(_patch(f"/mail/messages/{message_id}", data), default=str)


@mcp.tool()
def mail_batch_move(
    message_ids: list[str],
    destination_folder: str,
) -> str:
    """Move multiple messages to a folder in one operation.

    Args:
        message_ids: List of message IDs to move
        destination_folder: Destination folder ID or well-known name
    """
    return json.dumps(_post("/mail/batch/move", {
        "message_ids": message_ids,
        "destination_id": destination_folder,
    }), default=str)


@mcp.tool()
def mail_get_attachments(message_id: str) -> str:
    """List attachments for a message.

    Args:
        message_id: Message ID
    """
    return json.dumps(_get(f"/mail/messages/{message_id}/attachments"), default=str)


# ===========================================================================
# Calendar Tools
# ===========================================================================

@mcp.tool()
def calendar_list_calendars() -> str:
    """List all calendars the user has access to."""
    return json.dumps(_get("/calendar/calendars"), default=str)


@mcp.tool()
def calendar_list_events(
    calendar_id: str | None = None,
    top: int = 25,
    filter: str | None = None,
) -> str:
    """List upcoming calendar events.

    Args:
        calendar_id: Calendar ID (default: primary)
        top: Max events (default 25)
        filter: OData filter expression
    """
    params = {"top": top, "filter": filter}
    if calendar_id:
        params["calendar_id"] = calendar_id
    return json.dumps(_get("/calendar/events", params), default=str)


@mcp.tool()
def calendar_get_event(event_id: str) -> str:
    """Get a specific calendar event by ID.

    Args:
        event_id: Event ID
    """
    return json.dumps(_get(f"/calendar/events/{event_id}"), default=str)


@mcp.tool()
def calendar_view(
    start_datetime: str,
    end_datetime: str,
) -> str:
    """Get calendar events in a date range (calendar view).

    Args:
        start_datetime: Start in ISO 8601 (e.g. '2024-03-01T00:00:00')
        end_datetime: End in ISO 8601 (e.g. '2024-03-07T23:59:59')
    """
    return json.dumps(_get("/calendar/view", {
        "start_datetime": start_datetime,
        "end_datetime": end_datetime,
    }), default=str)


@mcp.tool()
def calendar_create_event(
    subject: str,
    start_datetime: str,
    end_datetime: str,
    body: str | None = None,
    location: str | None = None,
    attendees: list[str] | None = None,
    is_online_meeting: bool = False,
) -> str:
    """Create a calendar event.

    Args:
        subject: Event title
        start_datetime: Start time ISO 8601 (e.g. '2024-03-15T10:00:00')
        end_datetime: End time ISO 8601 (e.g. '2024-03-15T11:00:00')
        body: Event description (optional)
        location: Location (optional)
        attendees: List of attendee email addresses (optional)
        is_online_meeting: Create as Teams meeting (default False)
    """
    data = {
        "subject": subject,
        "start_datetime": start_datetime,
        "end_datetime": end_datetime,
        "is_online_meeting": is_online_meeting,
    }
    if body:
        data["body"] = body
    if location:
        data["location"] = location
    if attendees:
        data["attendees"] = attendees
    return json.dumps(_post("/calendar/events", data), default=str)


@mcp.tool()
def calendar_update_event(
    event_id: str,
    subject: str | None = None,
    start_datetime: str | None = None,
    end_datetime: str | None = None,
    body: str | None = None,
    location: str | None = None,
) -> str:
    """Update a calendar event. Only provided fields are changed.

    Args:
        event_id: Event ID
        subject: New title (optional)
        start_datetime: New start time (optional)
        end_datetime: New end time (optional)
        body: New description (optional)
        location: New location (optional)
    """
    data = {}
    if subject:
        data["subject"] = subject
    if start_datetime:
        data["start_datetime"] = start_datetime
    if end_datetime:
        data["end_datetime"] = end_datetime
    if body:
        data["body"] = body
    if location:
        data["location"] = location
    return json.dumps(_patch(f"/calendar/events/{event_id}", data), default=str)


@mcp.tool()
def calendar_delete_event(event_id: str) -> str:
    """Delete a calendar event.

    Args:
        event_id: Event ID
    """
    return json.dumps(_delete(f"/calendar/events/{event_id}"), default=str)


@mcp.tool()
def calendar_accept_event(event_id: str) -> str:
    """Accept a calendar event invitation.

    Args:
        event_id: Event ID
    """
    return json.dumps(_post(f"/calendar/events/{event_id}/accept"), default=str)


@mcp.tool()
def calendar_decline_event(event_id: str) -> str:
    """Decline a calendar event invitation.

    Args:
        event_id: Event ID
    """
    return json.dumps(_post(f"/calendar/events/{event_id}/decline"), default=str)


# ===========================================================================
# Files / OneDrive Tools
# ===========================================================================

@mcp.tool()
def files_search(
    q: str,
    drive_id: str | None = None,
    top: int = 25,
) -> str:
    """Search for files in OneDrive or SharePoint.

    Args:
        q: Search query
        drive_id: Drive ID to search in (default: user's OneDrive)
        top: Max results (default 25)
    """
    params = {"q": q, "top": top}
    if drive_id:
        params["drive_id"] = drive_id
    return json.dumps(_get("/files/search", params), default=str)


@mcp.tool()
def files_list_children(
    item_id: str,
) -> str:
    """List files and folders inside a OneDrive folder.

    Args:
        item_id: Folder item ID (use 'root' for root folder)
    """
    return json.dumps(_get(f"/files/items/{item_id}/children"), default=str)


@mcp.tool()
def files_get_item(item_id: str) -> str:
    """Get metadata for a file or folder.

    Args:
        item_id: Item ID
    """
    return json.dumps(_get(f"/files/items/{item_id}"), default=str)


# ===========================================================================
# SharePoint Tools
# ===========================================================================

@mcp.tool()
def sharepoint_list_drives() -> str:
    """List all SharePoint drives (document libraries) the user has access to."""
    return json.dumps(_get("/sharepoint/drives"), default=str)


@mcp.tool()
def sharepoint_list_children(
    item_id: str,
    drive_id: str | None = None,
    top: int = 100,
) -> str:
    """List files and folders in a SharePoint document library folder.

    Args:
        item_id: Folder item ID (use 'root' for root)
        drive_id: Drive ID (from sharepoint_list_drives). Required for non-default drives.
        top: Max results (default 100)
    """
    params = {"top": top}
    if drive_id:
        params["drive_id"] = drive_id
    return json.dumps(_get(f"/sharepoint/items/{item_id}/children", params), default=str)


@mcp.tool()
def sharepoint_search(
    q: str,
    drive_id: str | None = None,
    top: int = 25,
) -> str:
    """Search files in SharePoint.

    Args:
        q: Search query
        drive_id: Drive ID to search in
        top: Max results (default 25)
    """
    params = {"q": q, "top": top}
    if drive_id:
        params["drive_id"] = drive_id
    return json.dumps(_get("/sharepoint/search", params), default=str)


@mcp.tool()
def sharepoint_get_item(item_id: str) -> str:
    """Get metadata for a SharePoint file or folder.

    Args:
        item_id: Item ID
    """
    return json.dumps(_get(f"/sharepoint/items/{item_id}"), default=str)


# ===========================================================================
# Entry point
# ===========================================================================

MCP_AUTH_TOKEN = os.environ.get("MCP_AUTH_TOKEN", "")


def _make_auth_middleware(app):
    """Wrap a Starlette ASGI app with Bearer token auth."""
    from starlette.responses import Response as StarletteResponse

    async def middleware(scope, receive, send):
        if scope["type"] == "http" and MCP_AUTH_TOKEN:
            headers = dict(scope.get("headers", []))
            auth = headers.get(b"authorization", b"").decode()
            if auth != f"Bearer {MCP_AUTH_TOKEN}":
                response = StarletteResponse("Unauthorized", status_code=401)
                await response(scope, receive, send)
                return
        await app(scope, receive, send)

    return middleware


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="MS365 Access MCP Server")
    parser.add_argument(
        "--transport",
        choices=["stdio", "streamable-http"],
        default="stdio",
        help="MCP transport type (default: stdio)",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=8366,
        help="Port for streamable-http transport (default: 8366)",
    )
    args = parser.parse_args()

    if args.transport == "streamable-http":
        import uvicorn
        app = mcp.streamable_http_app()
        if MCP_AUTH_TOKEN:
            logger.info("MCP bearer auth enabled")
            app = _make_auth_middleware(app)
        uvicorn.run(app, host=_host, port=args.port)
    else:
        mcp.run(transport="stdio")
