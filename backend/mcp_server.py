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
from mcp.server.transport_security import TransportSecuritySettings

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

_host = os.environ.get("MCP_HOST", "127.0.0.1")
# Listen port = 8367 per the LaunchDaemon plist's --port arg. The hardcoded
# port below is only used for FastMCP's internal URL generation; uvicorn
# binds to args.port set below in main.
_PORT = 8367
_FQDN = "ft-deepak-m3-toronto.tailb4ec0f.ts.net"
mcp = FastMCP(
    "ms365-access",
    host=_host,
    port=_PORT,
    transport_security=TransportSecuritySettings(
        enable_dns_rebinding_protection=True,
        allowed_hosts=[
            "127.0.0.1:*", "localhost:*", "[::1]:*",
            f"{_FQDN}:{_PORT}",
        ],
        allowed_origins=[
            "http://127.0.0.1:*", "http://localhost:*", "http://[::1]:*",
            f"https://{_FQDN}:{_PORT}",
        ],
    ),
)

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


def _post(path: str, data: dict | None = None, params: dict | None = None) -> dict | list | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.post(path, json=data, params={k: v for k, v in (params or {}).items() if v is not None})
        r.raise_for_status()
        return r.json()


def _patch(path: str, data: dict | None = None, params: dict | None = None) -> dict | list | str:
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=30) as c:
        r = c.patch(path, json=data, params={k: v for k, v in (params or {}).items() if v is not None})
        r.raise_for_status()
        return r.json()


def _plain_to_html(text: str) -> str:
    """Convert plain text to HTML if it contains no HTML tags.

    Preserves newlines as <br> and escapes HTML entities so plain-text
    bodies render correctly in Outlook.
    """
    import html as _html
    import re
    if re.search(r"<[a-zA-Z][^>]*>", text):
        return text  # already contains HTML tags
    escaped = _html.escape(text)
    return escaped.replace("\n", "<br>\n")


def _put_binary(path: str, content: bytes, content_type: str = "application/octet-stream",
                 params: dict | None = None) -> dict | str:
    auth_headers = {"Authorization": f"Bearer {API_KEY}"}
    with httpx.Client(base_url=BASE_URL, headers=auth_headers, timeout=60) as c:
        r = c.put(
            path,
            params={k: v for k, v in (params or {}).items() if v is not None},
            files={"file": ("upload", content, content_type)},
        )
        r.raise_for_status()
        return r.json()


def _post_multipart(path: str, content: bytes, content_type: str,
                    fields: dict | None = None, params: dict | None = None) -> dict | str:
    """POST a file plus optional text form fields as multipart/form-data."""
    auth_headers = {"Authorization": f"Bearer {API_KEY}"}
    data = {k: v for k, v in (fields or {}).items() if v is not None}
    with httpx.Client(base_url=BASE_URL, headers=auth_headers, timeout=120) as c:
        r = c.post(
            path,
            params={k: v for k, v in (params or {}).items() if v is not None},
            files={"file": ("upload", content, content_type)},
            data=data,
        )
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
    include_body: bool = False,
    user: str | None = None,
) -> str:
    """List email messages, optionally filtered by folder or search.

    Args:
        folder: Folder ID or well-known name (inbox, archive, junkemail). Default: all messages.
        top: Max messages to return (default 25)
        search: Free-text search query
        filter: OData filter expression (e.g. "isRead eq false")
        include_body: Include truncated plain-text body (~2000 chars) alongside metadata (default False)
        user: UPN (e.g. "caroline@revivalgourmet.com") of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"top": top, "search": search, "filter": filter, "include_body": include_body}
    if folder:
        params["folder_id"] = folder
    if user:
        params["user"] = user
    return json.dumps(_get("/mail/messages", params), default=str)


@mcp.tool()
def mail_search(
    q: str,
    top: int = 25,
    user: str | None = None,
) -> str:
    """Search email messages by keyword.

    Args:
        q: Search query (searches subject, body, sender)
        top: Max results (default 25)
        user: UPN (e.g. "caroline@revivalgourmet.com") of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"q": q, "top": top}
    if user:
        params["user"] = user
    return json.dumps(_get("/mail/search", params), default=str)


@mcp.tool()
def mail_get_message(message_id: str, user: str | None = None) -> str:
    """Get a single email message by ID with full body.

    Args:
        message_id: Message ID
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"user": user} if user else None
    return json.dumps(_get(f"/mail/messages/{message_id}", params), default=str)


@mcp.tool()
def mail_get_threads(
    folder: str | None = None,
    top: int = 25,
    user: str | None = None,
) -> str:
    """List email threads (conversations) grouped by conversationId.

    Args:
        folder: Folder ID or well-known name. Default: all.
        top: Max threads (default 25)
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"top": top}
    if folder:
        params["folder_id"] = folder
    if user:
        params["user"] = user
    return json.dumps(_get("/mail/threads", params), default=str)


@mcp.tool()
def mail_create_draft(
    subject: str,
    body: str = "",
    to_recipients: list[str] | None = None,
    cc_recipients: list[str] | None = None,
    body_type: str = "HTML",
    user: str | None = None,
) -> str:
    """Create a new draft email in the Drafts folder (does not send it).

    For NEW emails only. To reply to an existing thread, use mail_create_reply_draft
    instead — it preserves threading and conversation context in Outlook.

    Plain text bodies are auto-converted to HTML (newlines become <br>) so
    formatting is preserved in Outlook. You can also pass raw HTML directly.

    Args:
        subject: Email subject
        body: Email body content (default empty). Plain text newlines are preserved.
        to_recipients: List of recipient email addresses (optional)
        cc_recipients: CC recipients (optional)
        body_type: 'HTML' or 'Text' (default 'HTML')
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    # Auto-convert plain text newlines to HTML so Outlook preserves formatting
    if body_type.upper() == "HTML" and body:
        body = _plain_to_html(body)
    data: dict = {"subject": subject, "body": body, "body_type": body_type}
    if to_recipients:
        data["to_recipients"] = to_recipients
    if cc_recipients:
        data["cc_recipients"] = cc_recipients
    params = {"user": user} if user else None
    return json.dumps(_post("/mail/drafts", data, params=params), default=str)


@mcp.tool()
def mail_create_reply_draft(
    message_id: str,
    comment: str = "",
    reply_all: bool = False,
    user: str | None = None,
) -> str:
    """Create a draft reply to a message (does not send it).

    This is the correct tool for replying to existing emails — it preserves
    the conversation thread in Outlook. Plain text newlines in the comment
    are auto-converted to HTML <br> tags.

    Args:
        message_id: ID of the message to reply to
        comment: Reply text to prepend to the quoted thread (default empty). Newlines preserved.
        reply_all: Create reply-all draft instead of reply (default False)
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    # Auto-convert plain text newlines to HTML so Outlook preserves formatting
    if comment:
        comment = _plain_to_html(comment)
    params = {}
    if reply_all:
        params["reply_all"] = "true"
    if user:
        params["user"] = user
    return json.dumps(
        _post(
            f"/mail/messages/{message_id}/draftReply",
            {"comment": comment} if comment else None,
            params=params or None,
        ),
        default=str,
    )


@mcp.tool()
def mail_send(
    subject: str,
    body: str,
    to_recipients: list[str],
    cc_recipients: list[str] | None = None,
    body_type: str = "HTML",
    user: str | None = None,
) -> str:
    """Send an email message.

    Args:
        subject: Email subject
        body: Email body content
        to_recipients: List of recipient email addresses
        cc_recipients: CC recipients (optional)
        body_type: 'HTML' or 'Text' (default 'HTML')
        user: UPN of a shared mailbox you have Send-As/Send-on-Behalf rights to. Default: your own mailbox.
    """
    data = {
        "subject": subject,
        "body": body,
        "to_recipients": to_recipients,
        "body_type": body_type,
    }
    if cc_recipients:
        data["cc_recipients"] = cc_recipients
    params = {"user": user} if user else None
    return json.dumps(_post("/mail/messages", data, params=params), default=str)


@mcp.tool()
def mail_reply(
    message_id: str,
    comment: str,
    user: str | None = None,
) -> str:
    """Reply to an email message.

    Args:
        message_id: ID of the message to reply to
        comment: Reply body text
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"user": user} if user else None
    return json.dumps(
        _post(f"/mail/messages/{message_id}/reply", {"comment": comment}, params=params),
        default=str,
    )


@mcp.tool()
def mail_move(
    message_id: str,
    destination_folder: str,
    user: str | None = None,
) -> str:
    """Move a message to a different folder.

    Args:
        message_id: Message ID
        destination_folder: Destination folder ID or well-known name (archive, junkemail, deleteditems)
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"user": user} if user else None
    return json.dumps(
        _post(
            f"/mail/messages/{message_id}/move",
            {"destination_folder_id": destination_folder},
            params=params,
        ),
        default=str,
    )


@mcp.tool()
def mail_update(
    message_id: str,
    is_read: bool | None = None,
    flag: str | None = None,
    body: str | None = None,
    body_type: str = "HTML",
    subject: str | None = None,
    to_recipients: list | None = None,
    cc_recipients: list | None = None,
    user: str | None = None,
) -> str:
    """Update message properties (mark read/unread, flag, body, subject, recipients).

    subject, to_recipients, and cc_recipients only work on draft messages.

    Args:
        message_id: Message ID
        is_read: Mark as read (True) or unread (False)
        flag: Flag status - 'flagged', 'complete', or 'notFlagged'
        body: New body content (HTML or plain text)
        body_type: 'HTML' or 'Text' (default 'HTML')
        subject: New subject (drafts only)
        to_recipients: New To recipients - email strings ("a@b.com"), RFC 5322 ("Name <a@b.com>"), or Graph API objects (drafts only)
        cc_recipients: New CC recipients - same formats as to_recipients (drafts only)
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    data = {}
    if is_read is not None:
        data["isRead"] = is_read
    if flag:
        data["flag"] = {"flagStatus": flag}
    if body is not None:
        data["body"] = body
        data["body_type"] = body_type
    if subject is not None:
        data["subject"] = subject
    if to_recipients is not None:
        data["to_recipients"] = to_recipients
    if cc_recipients is not None:
        data["cc_recipients"] = cc_recipients
    params = {"user": user} if user else None
    return json.dumps(_patch(f"/mail/messages/{message_id}", data, params=params), default=str)


@mcp.tool()
def mail_batch_move(
    message_ids: list[str],
    destination_folder: str,
    user: str | None = None,
) -> str:
    """Move multiple messages to a folder in one operation.

    Args:
        message_ids: List of message IDs to move
        destination_folder: Destination folder ID or well-known name
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"user": user} if user else None
    return json.dumps(
        _post(
            "/mail/batch/move",
            {"message_ids": message_ids, "destination_folder_id": destination_folder},
            params=params,
        ),
        default=str,
    )


@mcp.tool()
def mail_get_attachments(message_id: str, user: str | None = None) -> str:
    """List attachments for a message.

    Args:
        message_id: Message ID
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    params = {"user": user} if user else None
    return json.dumps(_get(f"/mail/messages/{message_id}/attachments", params), default=str)


@mcp.tool()
def mail_add_attachment(
    message_id: str,
    file_path: str,
    filename: str | None = None,
    content_type: str | None = None,
    user: str | None = None,
) -> str:
    """Add a file attachment to an email draft or message.

    Reads a local file, base64-encodes it, and attaches it to the message.
    Maximum file size ~3MB (MS Graph single-request attachment limit).

    Args:
        message_id: ID of the message or draft to attach to
        file_path: Absolute path to the local file to attach
        filename: Filename for the attachment (defaults to the local filename)
        content_type: MIME type (auto-detected from filename if not provided)
        user: UPN of a shared mailbox you have Full Access to. Default: your own mailbox.
    """
    import base64
    import mimetypes

    if not os.path.isfile(file_path):
        return json.dumps({"error": f"File not found: {file_path}"})

    file_size = os.path.getsize(file_path)
    if file_size > 3 * 1024 * 1024:
        return json.dumps({
            "error": f"File too large: {file_size} bytes. Maximum is 3MB for single-request attachments.",
            "size": file_size,
        })

    if not filename:
        filename = os.path.basename(file_path)

    if not content_type:
        content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"

    with open(file_path, "rb") as f:
        content_bytes = base64.b64encode(f.read()).decode("ascii")

    data = {
        "name": filename,
        "content_bytes": content_bytes,
        "content_type": content_type,
    }
    params = {"user": user} if user else None
    result = _post(f"/mail/messages/{message_id}/attachments", data, params=params)
    return json.dumps(result, default=str)


# ===========================================================================
# Calendar Tools
# ===========================================================================

@mcp.tool()
def calendar_list_calendars(user: str | None = None) -> str:
    """List all calendars the user has access to.

    Args:
        user: UPN of another user whose calendar list you have Full Access to. Default: your own.
    """
    params = {"user": user} if user else None
    return json.dumps(_get("/calendar/calendars", params), default=str)


@mcp.tool()
def calendar_list_events(
    calendar_id: str | None = None,
    top: int = 25,
    skip: int = 0,
    order_by: str = "start/dateTime desc",
    filter: str | None = None,
    user: str | None = None,
) -> str:
    """List calendar events from the event store (newest first by default).

    NOTE: This queries raw event records and does NOT expand recurring events —
    a weekly meeting appears as a single master record. For date-range queries
    that expand recurrences into individual occurrences, use calendar_view instead.

    Args:
        calendar_id: Calendar ID (default: primary)
        top: Max events (default 25)
        skip: Number of events to skip for pagination (default 0)
        order_by: OData $orderby expression (default "start/dateTime desc" — newest first). Use "start/dateTime" for oldest first.
        filter: OData filter expression
        user: UPN of another user whose calendar you have Full Access to. Default: your own.
    """
    params = {"top": top, "skip": skip, "order_by": order_by, "filter": filter}
    if calendar_id:
        params["calendar_id"] = calendar_id
    if user:
        params["user"] = user
    return json.dumps(_get("/calendar/events", params), default=str)


@mcp.tool()
def calendar_get_event(event_id: str, user: str | None = None) -> str:
    """Get a specific calendar event by ID.

    Args:
        event_id: Event ID
        user: UPN of another user whose calendar you have Full Access to. Default: your own.
    """
    params = {"user": user} if user else None
    return json.dumps(_get(f"/calendar/events/{event_id}", params), default=str)


@mcp.tool()
def calendar_view(
    start_datetime: str,
    end_datetime: str,
    calendar_id: str | None = None,
    top: int = 100,
    user: str | None = None,
) -> str:
    """Get calendar events in a date range — the preferred tool for upcoming events.

    Unlike calendar_list_events, this expands recurring events into individual
    occurrences within the requested window. Use this for "what's on my calendar
    today/this week" queries.

    Args:
        start_datetime: Start in ISO 8601 (e.g. '2026-02-18T00:00:00')
        end_datetime: End in ISO 8601 (e.g. '2026-02-18T23:59:59')
        calendar_id: Calendar ID (default: primary)
        top: Max events to return (default 100)
        user: UPN of another user whose calendar you have Full Access to. Default: your own.
    """
    params = {
        "start_datetime": start_datetime,
        "end_datetime": end_datetime,
        "top": top,
    }
    if calendar_id:
        params["calendar_id"] = calendar_id
    if user:
        params["user"] = user
    return json.dumps(_get("/calendar/view", params), default=str)


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
    user: str | None = None,
) -> str:
    """Search for files in OneDrive or SharePoint.

    Args:
        q: Search query
        drive_id: Drive ID to search in (default: user's OneDrive)
        top: Max results (default 25)
        user: UPN (e.g. "caroline@revivalgourmet.com") of another user whose OneDrive you have access to. Default: your own OneDrive. Ignored if drive_id is provided.
    """
    params = {"q": q, "top": top}
    if drive_id:
        params["drive_id"] = drive_id
    if user:
        params["user"] = user
    return json.dumps(_get("/files/search", params), default=str)


@mcp.tool()
def files_list_children(
    item_id: str,
    drive_id: str | None = None,
    user: str | None = None,
) -> str:
    """List files and folders inside a OneDrive folder.

    Args:
        item_id: Folder item ID (use 'root' for root folder)
        drive_id: Drive ID (default: user's OneDrive). Required for SharePoint drives.
        user: UPN of another user whose OneDrive folder you have access to. Default: your own. Ignored if drive_id is provided.
    """
    params = {}
    if drive_id:
        params["drive_id"] = drive_id
    if user:
        params["user"] = user
    return json.dumps(_get(f"/files/items/{item_id}/children", params or None), default=str)


@mcp.tool()
def files_get_item(item_id: str, drive_id: str | None = None, user: str | None = None) -> str:
    """Get metadata for a file or folder.

    Args:
        item_id: Item ID
        drive_id: Drive ID (default: user's OneDrive). Required for SharePoint drives.
        user: UPN of another user whose OneDrive item you have access to. Default: your own. Ignored if drive_id is provided.
    """
    params = {}
    if drive_id:
        params["drive_id"] = drive_id
    if user:
        params["user"] = user
    return json.dumps(_get(f"/files/items/{item_id}", params or None), default=str)


@mcp.tool()
def files_download_file(
    item_id: str,
    filename: str,
    drive_id: str | None = None,
    max_size_bytes: int = 10_485_760,
    user: str | None = None,
) -> str:
    """Download a OneDrive/SharePoint file (PDF, XLSX, DOCX, etc.) and save to a local temp file.

    The file is saved to /tmp/ms365-download-{item_id}-{safe_filename}.

    Args:
        item_id: File item ID
        filename: Original filename (for safe naming in /tmp)
        drive_id: Drive ID (required for SharePoint drives)
        max_size_bytes: Maximum file size in bytes (default 10MB)
        user: UPN of another user whose OneDrive file you have access to. Default: your own. Ignored if drive_id is provided.
    """
    import os
    import re

    # Sanitize filename
    safe_filename = re.sub(r'[/\\:\x00]', '_', os.path.basename(filename))
    if not safe_filename:
        safe_filename = "download"

    # Download raw bytes
    params = {}
    if drive_id:
        params["drive_id"] = drive_id
    if user:
        params["user"] = user
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=60) as c:
        r = c.get(f"/files/items/{item_id}/content", params=params)
        r.raise_for_status()
        data = r.content

    if len(data) > max_size_bytes:
        return json.dumps({
            "error": f"File too large: {len(data)} bytes (max {max_size_bytes})",
            "size": len(data),
        })

    dest_path = f"/tmp/ms365-download-{item_id}-{safe_filename}"
    with open(dest_path, "wb") as f:
        f.write(data)

    return json.dumps({
        "path": dest_path,
        "size": len(data),
        "filename": safe_filename,
    })


@mcp.tool()
def files_upload_file(
    local_path: str,
    parent_id: str,
    filename: str | None = None,
    drive_id: str | None = None,
) -> str:
    """Upload a local file to OneDrive or SharePoint.

    Args:
        local_path: Absolute path to the local file to upload
        parent_id: Parent folder item ID (get from sharepoint_list_children)
        filename: Destination filename (defaults to the local filename)
        drive_id: Drive ID (required for SharePoint drives)
    """
    import os
    import mimetypes

    if not os.path.isfile(local_path):
        return json.dumps({"error": f"File not found: {local_path}"})

    if not filename:
        filename = os.path.basename(local_path)

    content_type = mimetypes.guess_type(filename)[0] or "application/octet-stream"

    with open(local_path, "rb") as f:
        file_bytes = f.read()

    from urllib.parse import quote
    safe_name = quote(filename, safe="")
    auth_headers = {"Authorization": f"Bearer {API_KEY}"}
    with httpx.Client(base_url=BASE_URL, headers=auth_headers, timeout=60) as c:
        r = c.put(
            f"/files/items/{parent_id}:/{safe_name}:/content",
            params={k: v for k, v in {"drive_id": drive_id}.items() if v is not None},
            files={"file": (filename, file_bytes, content_type)},
        )
        r.raise_for_status()
        result = r.json()

    return json.dumps({
        "id": result.get("id"),
        "name": result.get("name"),
        "webUrl": result.get("webUrl"),
        "size": result.get("size"),
    })


@mcp.tool()
def files_replace_file(
    item_id: str,
    local_path: str,
    drive_id: str | None = None,
) -> str:
    """Replace the content of an existing OneDrive/SharePoint file (keeps the same item ID).

    OneDrive/SharePoint preserves version history, so the previous content is
    still accessible via the file's version history.

    Args:
        item_id: Item ID of the existing file to replace
        local_path: Absolute path to the local file with the new content
        drive_id: Drive ID (required for SharePoint drives)
    """
    import os
    import mimetypes

    if not os.path.isfile(local_path):
        return json.dumps({"error": f"File not found: {local_path}"})

    content_type = mimetypes.guess_type(local_path)[0] or "application/octet-stream"

    with open(local_path, "rb") as f:
        file_bytes = f.read()

    params = {}
    if drive_id:
        params["drive_id"] = drive_id

    result = _put_binary(
        f"/files/items/{item_id}/content",
        content=file_bytes,
        content_type=content_type,
        params=params,
    )

    return json.dumps({
        "id": result.get("id") if isinstance(result, dict) else None,
        "name": result.get("name") if isinstance(result, dict) else None,
        "webUrl": result.get("webUrl") if isinstance(result, dict) else None,
        "size": result.get("size") if isinstance(result, dict) else None,
    })


@mcp.tool()
def files_smart_update(
    item_id: str,
    local_path: str,
    drive_id: str | None = None,
    site_id: str | None = None,
    region_map: dict | None = None,
) -> str:
    """Replace a SharePoint/OneDrive .xlsx in place, with a live-edit fallback.

    Tries a normal whole-file replace first. If the file is open in Excel and the
    replace is blocked (423 Locked), it diffs the new workbook against the live
    one and surgically applies the changes it can reproduce via a Graph Excel
    session — values/formulas inside the declared `region_map` regions, plus
    worksheet add/delete/rename/reorder — leaving all formatting (including
    conditional formatting) untouched. If the change can't be reproduced live
    (a formatting/structural change, a change outside any region, or an exclusive
    lock), it returns `deferred` so you can ask the user to close the file and
    retry a clean replace.

    Args:
        item_id: Item ID of the existing .xlsx to update
        local_path: Absolute path to the freshly built local .xlsx
        drive_id: Drive ID (required for SharePoint drives)
        site_id: SharePoint site ID (alternative to drive_id; omit for OneDrive)
        region_map: Optional per-tab data regions safe to overwrite, e.g.
            {"AP payments": {"data": "A2:U200"}}. Without it, most diffs defer
            (safe default) — the map is what enables live-editing.

    Returns JSON {"mode", "ranges_written", "reason"} where mode is one of
    `replaced` | `live-edited` | `deferred`.
    """
    import os

    if not os.path.isfile(local_path):
        return json.dumps({"error": f"File not found: {local_path}"})

    with open(local_path, "rb") as f:
        file_bytes = f.read()

    params = {}
    if drive_id:
        params["drive_id"] = drive_id
    if site_id:
        params["site_id"] = site_id

    fields = {}
    if region_map is not None:
        fields["region_map"] = json.dumps(region_map)

    content_type = (
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
    result = _post_multipart(
        f"/files/items/{item_id}/smart-update",
        content=file_bytes,
        content_type=content_type,
        fields=fields,
        params=params,
    )
    return json.dumps(result, default=str)


# ===========================================================================
# SharePoint Tools
# ===========================================================================

@mcp.tool()
def sharepoint_resolve_site(host_path: str) -> str:
    """Resolve a SharePoint site to get its site ID.

    Args:
        host_path: SharePoint hostname/path, e.g. 'revivalgourmet.sharepoint.com/sites/Finance'
                   or 'revivalgourmet.sharepoint.com' for root site.
    """
    return json.dumps(_get(f"/sharepoint/sites/{host_path}"), default=str)


@mcp.tool()
def sharepoint_list_drives(site_id: str) -> str:
    """List all SharePoint drives (document libraries) for a site.

    Args:
        site_id: SharePoint site ID. Get this from sharepoint_resolve_site
                 (e.g. 'revivalgourmet.sharepoint.com/sites/Finance').
    """
    return json.dumps(_get("/sharepoint/drives", {"site_id": site_id}), default=str)


@mcp.tool()
def sharepoint_list_children(
    item_id: str,
    site_id: str,
    top: int = 100,
) -> str:
    """List files and folders in a SharePoint document library folder.

    Args:
        item_id: Folder item ID (use 'root' for root)
        site_id: SharePoint site ID (from sharepoint_resolve_site)
        top: Max results (default 100)
    """
    params = {"site_id": site_id, "top": top}
    return json.dumps(_get(f"/sharepoint/items/{item_id}/children", params), default=str)


@mcp.tool()
def sharepoint_search(
    q: str,
    site_id: str,
    top: int = 25,
) -> str:
    """Search files in SharePoint.

    Args:
        q: Search query
        site_id: SharePoint site ID (from sharepoint_resolve_site)
        top: Max results (default 25)
    """
    params = {"q": q, "site_id": site_id, "top": top}
    return json.dumps(_get("/sharepoint/search", params), default=str)


@mcp.tool()
def sharepoint_get_item(item_id: str, site_id: str) -> str:
    """Get metadata for a SharePoint file or folder.

    Args:
        item_id: Item ID
        site_id: SharePoint site ID (from sharepoint_resolve_site)
    """
    return json.dumps(_get(f"/sharepoint/items/{item_id}", {"site_id": site_id}), default=str)


@mcp.tool()
def sharepoint_rename_item(
    item_id: str,
    site_id: str,
    new_name: str,
) -> str:
    """Rename a file or folder in SharePoint.

    Args:
        item_id: Item ID (from sharepoint_list_children or sharepoint_search)
        site_id: SharePoint site ID (from sharepoint_resolve_site)
        new_name: New name for the file or folder (include extension)
    """
    result = _patch(
        f"/sharepoint/items/{item_id}",
        data={"name": new_name},
        params={"site_id": site_id},
    )
    return json.dumps({
        "id": result.get("id") if isinstance(result, dict) else None,
        "name": result.get("name") if isinstance(result, dict) else None,
        "webUrl": result.get("webUrl") if isinstance(result, dict) else None,
    }, default=str)


@mcp.tool()
def sharepoint_move_item(
    item_id: str,
    site_id: str,
    destination_folder_id: str,
) -> str:
    """Move a file or folder to a different folder in SharePoint.

    Args:
        item_id: Item ID of the file/folder to move (from sharepoint_list_children or sharepoint_search)
        site_id: SharePoint site ID (from sharepoint_resolve_site)
        destination_folder_id: Item ID of the destination folder
    """
    result = _patch(
        f"/sharepoint/items/{item_id}/move",
        data={"destination_folder_id": destination_folder_id},
        params={"site_id": site_id},
    )
    return json.dumps({
        "id": result.get("id") if isinstance(result, dict) else None,
        "name": result.get("name") if isinstance(result, dict) else None,
        "webUrl": result.get("webUrl") if isinstance(result, dict) else None,
        "parentReference": result.get("parentReference") if isinstance(result, dict) else None,
    }, default=str)


@mcp.tool()
def sharepoint_resolve_url(url: str) -> str:
    """Resolve a SharePoint sharing URL to site_id and item_id.

    Handles sharing links like https://contoso.sharepoint.com/:p:/s/SiteName/EaBC123...
    Returns item metadata including site_id and item_id for use with other SharePoint tools.

    Args:
        url: SharePoint sharing URL or document URL
    """
    return json.dumps(_get("/sharepoint/resolve", {"url": url}), default=str)


@mcp.tool()
def sharepoint_download_from_url(
    url: str,
    filename: str = "",
    max_size_bytes: int = 10_485_760,
) -> str:
    """Download a SharePoint file directly from a sharing URL.

    Resolves the sharing link to site_id/item_id, then downloads the file to /tmp.

    Args:
        url: SharePoint sharing URL
        filename: Optional filename override (auto-detected from metadata if empty)
        max_size_bytes: Maximum file size in bytes (default 10MB)
    """
    import os
    import re

    # Step 1: Resolve URL to item metadata
    resolved = _get("/sharepoint/resolve", {"url": url})
    if not resolved or "error" in resolved:
        return json.dumps({"error": f"Could not resolve URL: {resolved}"})

    item_id = resolved.get("item_id")
    site_id = resolved.get("site_id")
    if not item_id or not site_id:
        return json.dumps({"error": "Resolved but missing item_id or site_id", "resolved": resolved})

    # Step 2: Determine filename
    if not filename:
        item = resolved.get("item", {})
        filename = item.get("name", f"download-{item_id}")

    safe_filename = re.sub(r'[/\\:\x00]', '_', os.path.basename(filename))
    if not safe_filename:
        safe_filename = "download"

    # Step 3: Download via SharePoint endpoint (uses site-scoped drive path)
    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=60) as c:
        r = c.get(f"/sharepoint/items/{item_id}/content", params={"site_id": site_id})
        r.raise_for_status()
        data = r.content

    if len(data) > max_size_bytes:
        return json.dumps({"error": f"File too large: {len(data)} bytes (max {max_size_bytes})", "size": len(data)})

    dest_path = f"/tmp/ms365-download-{item_id}-{safe_filename}"
    with open(dest_path, "wb") as f:
        f.write(data)

    return json.dumps({"path": dest_path, "size": len(data), "filename": safe_filename, "item_id": item_id, "site_id": site_id})


@mcp.tool()
def sharepoint_list_versions(item_id: str, site_id: str, top: int = 50) -> str:
    """List the version history for a SharePoint file.

    Each version has an id, lastModifiedDateTime, size, and lastModifiedBy.
    Use the version id with sharepoint_download_version to fetch that historical content.

    Args:
        item_id: Item ID (from sharepoint_search or sharepoint_list_children)
        site_id: SharePoint site ID (from sharepoint_resolve_url or sharepoint_resolve_site)
        top: Max versions to return (default 50)
    """
    return json.dumps(
        _get(f"/sharepoint/items/{item_id}/versions", {"site_id": site_id, "top": top}),
        default=str,
    )


@mcp.tool()
def sharepoint_download_version(
    item_id: str,
    version_id: str,
    site_id: str,
    filename: str = "",
    max_size_bytes: int = 10_485_760,
) -> str:
    """Download a specific historical version of a SharePoint file to /tmp.

    Use sharepoint_list_versions to find the version_id.

    Args:
        item_id: Item ID (from sharepoint_search or sharepoint_list_children)
        version_id: Version ID (from sharepoint_list_versions)
        site_id: SharePoint site ID (from sharepoint_resolve_url or sharepoint_resolve_site)
        filename: Optional filename override
        max_size_bytes: Maximum file size in bytes (default 10MB)
    """
    import os
    import re

    safe_filename = re.sub(r'[/\\:\x00]', '_', os.path.basename(filename))
    if not safe_filename:
        safe_filename = "download"

    with httpx.Client(base_url=BASE_URL, headers=_headers(), timeout=60) as c:
        r = c.get(
            f"/sharepoint/items/{item_id}/versions/{version_id}/content",
            params={"site_id": site_id},
        )
        r.raise_for_status()
        data = r.content

    if len(data) > max_size_bytes:
        return json.dumps({"error": f"File too large: {len(data)} bytes (max {max_size_bytes})", "size": len(data)})

    dest_path = f"/tmp/ms365-version-{item_id}-{version_id}-{safe_filename}"
    with open(dest_path, "wb") as f:
        f.write(data)

    return json.dumps({"path": dest_path, "size": len(data), "version_id": version_id})


# ===========================================================================
# Workbook (Excel) Tools — live, co-authoring-safe cell/range/table writes
# ===========================================================================
#
# These edit an .xlsx at the cell level via the Graph Excel API, so they merge
# safely into a workbook that is open on other machines (unlike a whole-file
# replace via files_replace_file / sharepoint upload, which conflicts with open
# editors). Writes to one workbook must be SEQUENTIAL.
#
# For a single edit, just call workbook_update_range / workbook_add_table_row
# with no session_id — the server opens and closes a persistent session around
# the write for you (a clean one-call in/out). For several edits to the same
# workbook, call workbook_create_session once, pass the returned session_id to
# each write, then workbook_close_session. site_id is required
# for SharePoint files (from sharepoint_resolve_url / sharepoint_resolve_site)
# and omitted for the signed-in user's OneDrive.


@mcp.tool()
def workbook_check_lock(item_id: str, site_id: str = "") -> str:
    """Check whether an Excel file is checked out (blocks writes).

    Co-authoring (multiple people editing) is NOT a checkout and is safe to
    write to. Use this before a session of edits.

    Args:
        item_id: Excel file item ID (from sharepoint_resolve_url / search)
        site_id: SharePoint site ID; omit for OneDrive
    """
    params = {"site_id": site_id or None}
    return json.dumps(_get(f"/workbook/items/{item_id}/lock-state", params), default=str)


@mcp.tool()
def workbook_create_session(item_id: str, site_id: str = "", persist: bool = True) -> str:
    """Open an Excel workbook session for a batch of edits; returns a session id.

    Pass the returned 'id' as session_id to subsequent workbook_* calls, then
    call workbook_close_session. A persistent session merges writes into the
    live co-authoring session. Returns HTTP 409 if the file is checked out.

    Args:
        item_id: Excel file item ID
        site_id: SharePoint site ID; omit for OneDrive
        persist: Persist changes to the file (default True)
    """
    params = {"site_id": site_id or None, "persist": persist}
    return json.dumps(_post(f"/workbook/items/{item_id}/session", params=params), default=str)


@mcp.tool()
def workbook_close_session(item_id: str, session_id: str, site_id: str = "") -> str:
    """Close an Excel workbook session opened with workbook_create_session.

    Args:
        item_id: Excel file item ID
        session_id: Session id from workbook_create_session
        site_id: SharePoint site ID; omit for OneDrive
    """
    qs = f"session_id={session_id}"
    if site_id:
        qs += f"&site_id={site_id}"
    return json.dumps(_delete(f"/workbook/items/{item_id}/session?{qs}"), default=str)


@mcp.tool()
def workbook_list_worksheets(item_id: str, site_id: str = "", session_id: str = "") -> str:
    """List the worksheets (tabs) in an Excel workbook.

    Args:
        item_id: Excel file item ID
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    return json.dumps(_get(f"/workbook/items/{item_id}/worksheets", params), default=str)


@mcp.tool()
def workbook_list_tables(item_id: str, site_id: str = "", session_id: str = "") -> str:
    """List the named tables in an Excel workbook.

    Args:
        item_id: Excel file item ID
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    return json.dumps(_get(f"/workbook/items/{item_id}/tables", params), default=str)


@mcp.tool()
def workbook_get_range(
    item_id: str,
    sheet: str,
    address: str,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Read a cell range from an Excel worksheet.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name, e.g. 'Sheet1'
        address: Range in A1 notation, e.g. 'A1:C10'
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {
        "sheet": sheet,
        "address": address,
        "site_id": site_id or None,
        "session_id": session_id or None,
    }
    return json.dumps(_get(f"/workbook/items/{item_id}/range", params), default=str)


@mcp.tool()
def workbook_update_range(
    item_id: str,
    sheet: str,
    address: str,
    values: list,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Write values into a cell range of a live Excel worksheet.

    Co-authoring-safe: merges into the file even while it is open elsewhere.
    `values` is a 2D array matching the address dimensions; formulas are
    strings starting with '='. Returns HTTP 409 if the file is checked out.
    With no session_id the server opens/closes a session for you (clean in/out).

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name, e.g. 'Sheet1'
        address: Range in A1 notation, e.g. 'A1:B2'
        values: Rows of cell values, e.g. [["Total", 42], ["=A1*2", 84]]
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional; pass one from workbook_create_session for batches
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet, "address": address, "values": values}
    return json.dumps(
        _patch(f"/workbook/items/{item_id}/range", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_add_table_row(
    item_id: str,
    table: str,
    values: list,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Append one or more rows to a named table in a live Excel workbook.

    Args:
        item_id: Excel file item ID
        table: Table name or id (from workbook_list_tables)
        values: Rows to append, e.g. [["east", "pear", 4]]
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"values": values}
    return json.dumps(
        _post(f"/workbook/items/{item_id}/tables/{table}/rows", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_get_worksheet(item_id: str, sheet: str, site_id: str = "", session_id: str = "") -> str:
    """Get a single worksheet's metadata (id, name, position, visibility).

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"sheet": sheet, "site_id": site_id or None, "session_id": session_id or None}
    return json.dumps(_get(f"/workbook/items/{item_id}/worksheet", params), default=str)


@mcp.tool()
def workbook_get_used_range(
    item_id: str,
    sheet: str,
    values_only: bool = False,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Get the used range of a worksheet (smallest range covering all data).

    Handy before reading/clearing to learn the extent of the data without
    guessing an address.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name, e.g. 'Sheet1'
        values_only: Exclude formatting-only cells, bounding to cells with values
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {
        "sheet": sheet,
        "values_only": values_only,
        "site_id": site_id or None,
        "session_id": session_id or None,
    }
    return json.dumps(_get(f"/workbook/items/{item_id}/used-range", params), default=str)


@mcp.tool()
def workbook_add_worksheet(item_id: str, name: str = "", site_id: str = "", session_id: str = "") -> str:
    """Add a new worksheet (tab) to a live Excel workbook.

    Co-authoring-safe. With no session_id the server opens/closes a session for
    you. Returns the new worksheet (including its assigned name and position).

    Args:
        item_id: Excel file item ID
        name: Name for the new worksheet; if omitted, Excel assigns a default
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional; pass one from workbook_create_session for batches
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"name": name} if name else {}
    return json.dumps(
        _post(f"/workbook/items/{item_id}/worksheets", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_delete_worksheet(item_id: str, sheet: str, site_id: str = "", session_id: str = "") -> str:
    """Delete a worksheet (tab) from a live Excel workbook.

    Co-authoring-safe. Deleting the only/last visible sheet will fail per Excel
    rules. With no session_id the server opens/closes a session for you.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id (from workbook_list_worksheets)
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    from urllib.parse import quote
    qs = f"sheet={quote(sheet)}"
    if site_id:
        qs += f"&site_id={site_id}"
    if session_id:
        qs += f"&session_id={session_id}"
    return json.dumps(_delete(f"/workbook/items/{item_id}/worksheet?{qs}"), default=str)


@mcp.tool()
def workbook_rename_worksheet(
    item_id: str,
    sheet: str,
    new_name: str,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Rename a worksheet in a live Excel workbook.

    Args:
        item_id: Excel file item ID
        sheet: Current worksheet name or id
        new_name: New worksheet name
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet, "name": new_name}
    return json.dumps(
        _patch(f"/workbook/items/{item_id}/worksheet", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_reorder_worksheet(
    item_id: str,
    sheet: str,
    position: int,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Move a worksheet to a new tab position in a live Excel workbook.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id
        position: 0-based target position (0 = first tab)
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet, "position": position}
    return json.dumps(
        _patch(f"/workbook/items/{item_id}/worksheet", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_update_worksheet(
    item_id: str,
    sheet: str,
    name: str = "",
    position: int | None = None,
    visibility: str = "",
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Update a worksheet's properties in one call (rename + reorder + show/hide).

    Generic counterpart to workbook_rename_worksheet / workbook_reorder_worksheet;
    use it to set visibility or change several properties at once. At least one
    of name / position / visibility must be provided.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id
        name: New name (omit to leave unchanged)
        position: 0-based tab position (omit to leave unchanged)
        visibility: 'Visible', 'Hidden', or 'VeryHidden' (omit to leave unchanged)
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body: dict = {"sheet": sheet}
    if name:
        body["name"] = name
    if position is not None:
        body["position"] = position
    if visibility:
        body["visibility"] = visibility
    return json.dumps(
        _patch(f"/workbook/items/{item_id}/worksheet", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_protect_worksheet(
    item_id: str,
    sheet: str,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Protect a worksheet (lock it against edits) in a live Excel workbook.

    Applies default protection. Use workbook_unprotect_worksheet to remove it.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet}
    return json.dumps(
        _post(f"/workbook/items/{item_id}/worksheet/protect", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_unprotect_worksheet(
    item_id: str,
    sheet: str,
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Remove protection from a worksheet in a live Excel workbook.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name or id
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet}
    return json.dumps(
        _post(f"/workbook/items/{item_id}/worksheet/unprotect", data=body, params=params),
        default=str,
    )


@mcp.tool()
def workbook_clear_range(
    item_id: str,
    sheet: str,
    address: str,
    apply_to: str = "All",
    site_id: str = "",
    session_id: str = "",
) -> str:
    """Clear cells in a range of a live Excel worksheet.

    Co-authoring-safe. With no session_id the server opens/closes a session for you.

    Args:
        item_id: Excel file item ID
        sheet: Worksheet name, e.g. 'Sheet1'
        address: Range in A1 notation, e.g. 'A1:C10'
        apply_to: What to clear — 'All', 'Formats', or 'Contents' (default 'All')
        site_id: SharePoint site ID; omit for OneDrive
        session_id: Optional workbook session id
    """
    params = {"site_id": site_id or None, "session_id": session_id or None}
    body = {"sheet": sheet, "address": address, "apply_to": apply_to}
    return json.dumps(
        _post(f"/workbook/items/{item_id}/range/clear", data=body, params=params),
        default=str,
    )


# ===========================================================================
# Contacts Tools
# ===========================================================================

@mcp.tool()
def contacts_list(
    top: int = 100,
    skip: int = 0,
    search: str | None = None,
) -> str:
    """List MS365 contacts. Optionally search by name or keyword.

    Args:
        top: Max contacts to return (default 100)
        skip: Number to skip for pagination (default 0)
        search: Free-text search query (optional)
    """
    params = {"top": top, "skip": skip}
    if search:
        params["search"] = search
    return json.dumps(_get("/contacts/", params), default=str)


@mcp.tool()
def contacts_get(contact_id: str) -> str:
    """Get a single MS365 contact by ID.

    Args:
        contact_id: Contact ID
    """
    return json.dumps(_get(f"/contacts/{contact_id}"), default=str)


@mcp.tool()
def contacts_create(
    name: str,
    email: str | None = None,
    phone: str | None = None,
    organization: str | None = None,
    title: str | None = None,
    notes: str | None = None,
) -> str:
    """Create a new MS365 contact.

    Args:
        name: Full name (given + surname)
        email: Email address (optional)
        phone: Mobile phone number (optional)
        organization: Company/organization (optional)
        title: Job title (optional)
        notes: Personal notes (optional)
    """
    data = {"name": name}
    if email:
        data["email"] = email
    if phone:
        data["phone"] = phone
    if organization:
        data["organization"] = organization
    if title:
        data["title"] = title
    if notes:
        data["notes"] = notes
    return json.dumps(_post("/contacts/", data), default=str)


@mcp.tool()
def contacts_update(
    contact_id: str,
    name: str | None = None,
    email: str | None = None,
    phone: str | None = None,
    organization: str | None = None,
    title: str | None = None,
    notes: str | None = None,
) -> str:
    """Update an existing MS365 contact. Only provided fields are changed.

    Args:
        contact_id: Contact ID
        name: New full name (optional)
        email: New email address (optional)
        phone: New phone number (optional)
        organization: New company name (optional)
        title: New job title (optional)
        notes: New personal notes (optional)
    """
    data = {}
    if name is not None:
        data["name"] = name
    if email is not None:
        data["email"] = email
    if phone is not None:
        data["phone"] = phone
    if organization is not None:
        data["organization"] = organization
    if title is not None:
        data["title"] = title
    if notes is not None:
        data["notes"] = notes
    return json.dumps(_patch(f"/contacts/{contact_id}", data), default=str)


@mcp.tool()
def contacts_delete(contact_id: str) -> str:
    """Delete an MS365 contact.

    Args:
        contact_id: Contact ID
    """
    return json.dumps(_delete(f"/contacts/{contact_id}"), default=str)


@mcp.tool()
def contacts_search_by_email(email: str) -> str:
    """Search for an MS365 contact by exact email address.

    Args:
        email: Email address to search for
    """
    return json.dumps(_get(f"/contacts/by-email/{email}"), default=str)


# ===========================================================================
# Power BI Tools
# ===========================================================================

@mcp.tool()
def powerbi_list_workspaces() -> str:
    """List all Power BI workspaces the user has access to.

    Returns workspace names, IDs, and capacity info.
    """
    return json.dumps(_get("/powerbi/workspaces"), default=str)


@mcp.tool()
def powerbi_list_datasets(workspace_id: str) -> str:
    """List datasets in a Power BI workspace.

    Args:
        workspace_id: Workspace (group) ID
    """
    return json.dumps(_get(f"/powerbi/workspaces/{workspace_id}/datasets"), default=str)


@mcp.tool()
def powerbi_list_tables(workspace_id: str, dataset_id: str) -> str:
    """List tables in a Power BI dataset.

    For standard (non-push) datasets, uses DAX INFO.TABLES() to discover tables.

    Args:
        workspace_id: Workspace (group) ID
        dataset_id: Dataset ID
    """
    return json.dumps(
        _get(f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/tables"),
        default=str,
    )


@mcp.tool()
def powerbi_query(workspace_id: str, dataset_id: str, dax_query: str) -> str:
    """Execute a DAX query against a Power BI dataset.

    Requires Power BI Premium or Premium Per User capacity.
    Returns columns and rows from the first result table.

    Example DAX queries:
        EVALUATE TOPN(10, 'Sales', 'Sales'[Amount], DESC)
        EVALUATE SUMMARIZECOLUMNS('Product'[Category], "Total", SUM('Sales'[Amount]))
        EVALUATE INFO.TABLES()

    Args:
        workspace_id: Workspace (group) ID
        dataset_id: Dataset ID
        dax_query: DAX query string (must start with EVALUATE, DEFINE, or a valid DAX statement)
    """
    return json.dumps(
        _post(
            f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/query",
            {"dax_query": dax_query},
        ),
        default=str,
    )


@mcp.tool()
def powerbi_list_reports(workspace_id: str) -> str:
    """List reports in a Power BI workspace.

    Args:
        workspace_id: Workspace (group) ID
    """
    return json.dumps(_get(f"/powerbi/workspaces/{workspace_id}/reports"), default=str)


@mcp.tool()
def powerbi_refresh_dataset(workspace_id: str, dataset_id: str) -> str:
    """Trigger a Power BI dataset refresh. Returns immediately — poll powerbi_list_refreshes for status.

    Recommended polling interval: 30-60 seconds.
    Power BI Pro: 8 refreshes/day max. Premium: 48/day.

    Args:
        workspace_id: Workspace (group) ID
        dataset_id: Dataset ID
    """
    return json.dumps(
        _post(
            f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/refreshes",
            {},
        ),
        default=str,
    )


@mcp.tool()
def powerbi_list_refreshes(workspace_id: str, dataset_id: str, top: int = 10) -> str:
    """List recent Power BI dataset refresh history (status, start/end time, type).

    Use after powerbi_refresh_dataset to check if a refresh completed or failed.
    Status values: Unknown, Completed, Failed, Cancelled, Disabled.

    Args:
        workspace_id: Workspace (group) ID
        dataset_id: Dataset ID
        top: Number of recent refreshes to return (default 10, max 100)
    """
    return json.dumps(
        _get(
            f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/refreshes",
            {"top": top},
        ),
        default=str,
    )


# ===========================================================================
# Email Linter Tool
# ===========================================================================

@mcp.tool()
def lint_email_addresses(addresses: list[str]) -> str:
    """Lint email addresses: enrich with display names from entity graph and Google Contacts.

    Call this before drafting emails to ensure addresses have correct display names.
    Returns each address with a warning: "enriched" (name found and applied),
    null (already correct), or "unknown" (no match found).

    Args:
        addresses: List of email address strings (bare emails or "Name <email>" format)
    """
    from lib.email_linter import lint_addresses_sync, load_entity_index

    entity_index = load_entity_index()

    gsuite_api_key = os.environ.get("GSUITE_API_KEY", "")

    def contact_lookup(email: str) -> str | None:
        try:
            r = httpx.get(
                "http://127.0.0.1:8001/contacts/by-email/" + email,
                params={"account": "personal"},
                headers={"Authorization": f"Bearer {gsuite_api_key}"},
                timeout=5,
            )
            if r.status_code == 200:
                data = r.json()
                return data.get("name") if data.get("found") else None
        except Exception as e:
            logger.warning("Contact lookup failed for %s: %s", email, e)
        return None

    result = lint_addresses_sync(addresses, entity_index, contact_lookup)
    return json.dumps(result, default=str)


# ===========================================================================
# Entry point
# ===========================================================================

MCP_AUTH_TOKEN = os.environ.get("MCP_AUTH_TOKEN", "")


def _make_auth_middleware(app):
    """Wrap a Starlette ASGI app with Bearer token auth."""
    from starlette.responses import Response as StarletteResponse

    async def middleware(scope, receive, send):
        if scope["type"] == "http":
            if not MCP_AUTH_TOKEN:
                response = StarletteResponse(
                    "MCP_AUTH_TOKEN not configured", status_code=503
                )
                await response(scope, receive, send)
                return
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
