"""Audit logging for sensitive operations."""

import logging
import json
from datetime import datetime
from pathlib import Path
from typing import Optional
from functools import wraps

from app.config import get_settings

settings = get_settings()

# Create audit logger with file handler
audit_logger = logging.getLogger("audit")
audit_logger.setLevel(logging.INFO)
audit_logger.propagate = False  # Don't propagate to root logger

# Ensure data directory exists
data_dir = Path("data")
data_dir.mkdir(exist_ok=True)

# File handler for audit log
file_handler = logging.FileHandler(data_dir / "audit.log")
file_handler.setLevel(logging.INFO)
file_handler.setFormatter(
    logging.Formatter("%(asctime)s - %(message)s", datefmt="%Y-%m-%d %H:%M:%S")
)
audit_logger.addHandler(file_handler)


def log_event(
    event_type: str,
    action: str,
    email: Optional[str] = None,
    details: Optional[dict] = None,
    success: bool = True,
):
    """Log an audit event.

    Args:
        event_type: Category of event (auth, mail, calendar, files)
        action: Specific action taken
        email: User email if available
        details: Additional context
        success: Whether the operation succeeded
    """
    entry = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "type": event_type,
        "action": action,
        "email": email,
        "success": success,
    }
    if details:
        entry["details"] = details

    audit_logger.info(json.dumps(entry))


# Auth events
def log_login_attempt(email: str, success: bool, error: Optional[str] = None):
    log_event(
        "auth",
        "login",
        email=email,
        success=success,
        details={"error": error} if error else None,
    )


def log_logout(email: str):
    log_event("auth", "logout", email=email)


def log_token_refresh(email: str, success: bool, error: Optional[str] = None):
    log_event(
        "auth",
        "token_refresh",
        email=email,
        success=success,
        details={"error": error} if error else None,
    )


# Mail events
def log_mail_send(email: str, to_recipients: list[str], subject: str):
    log_event(
        "mail",
        "send",
        email=email,
        details={"to": to_recipients, "subject": subject[:100]},
    )


def log_mail_delete(email: str, message_id: str):
    log_event("mail", "delete", email=email, details={"message_id": message_id})


def log_mail_batch_delete(email: str, message_ids: list[str]):
    log_event(
        "mail",
        "batch_delete",
        email=email,
        details={"count": len(message_ids), "message_ids": message_ids[:10]},
    )


def log_mail_move(email: str, message_id: str, destination: str):
    log_event(
        "mail",
        "move",
        email=email,
        details={"message_id": message_id, "destination": destination},
    )


# Calendar events
def log_calendar_create(email: str, subject: str, start: str, attendees: list[str]):
    log_event(
        "calendar",
        "create_event",
        email=email,
        details={
            "subject": subject[:100],
            "start": start,
            "attendees": attendees[:10],
        },
    )


def log_calendar_delete(email: str, event_id: str):
    log_event("calendar", "delete_event", email=email, details={"event_id": event_id})


def log_calendar_update(email: str, event_id: str, changes: list[str]):
    log_event(
        "calendar",
        "update_event",
        email=email,
        details={"event_id": event_id, "changed_fields": changes},
    )


# File events
def log_file_upload(email: str, filename: str, parent_id: str):
    log_event(
        "files",
        "upload",
        email=email,
        details={"filename": filename, "parent_id": parent_id},
    )


def log_file_delete(email: str, item_id: str):
    log_event("files", "delete", email=email, details={"item_id": item_id})


def log_file_download(email: str, item_id: str):
    log_event("files", "download", email=email, details={"item_id": item_id})
