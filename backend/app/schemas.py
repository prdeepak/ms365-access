from pydantic import BaseModel, EmailStr, Field
from datetime import datetime
from typing import Optional


# Auth schemas
class AuthStatus(BaseModel):
    authenticated: bool
    email: Optional[str] = None
    expires_at: Optional[datetime] = None


# Mail schemas
class MailFolder(BaseModel):
    id: str
    display_name: str
    parent_folder_id: Optional[str] = None
    child_folder_count: int = 0
    unread_item_count: int = 0
    total_item_count: int = 0


class EmailAddress(BaseModel):
    name: Optional[str] = None
    address: str


class Recipient(BaseModel):
    email_address: EmailAddress = Field(alias="emailAddress")

    class Config:
        populate_by_name = True


class MailMessage(BaseModel):
    id: str
    subject: Optional[str] = None
    body_preview: Optional[str] = Field(None, alias="bodyPreview")
    from_: Optional[Recipient] = Field(None, alias="from")
    to_recipients: list[Recipient] = Field(default_factory=list, alias="toRecipients")
    cc_recipients: list[Recipient] = Field(default_factory=list, alias="ccRecipients")
    received_datetime: Optional[datetime] = Field(None, alias="receivedDateTime")
    sent_datetime: Optional[datetime] = Field(None, alias="sentDateTime")
    is_read: bool = Field(False, alias="isRead")
    is_draft: bool = Field(False, alias="isDraft")
    has_attachments: bool = Field(False, alias="hasAttachments")
    importance: str = "normal"
    flag: Optional[dict] = None

    class Config:
        populate_by_name = True


class MailMessageDetail(MailMessage):
    body: Optional[dict] = None


class CreateDraftRequest(BaseModel):
    subject: str
    body: str = ""
    body_type: str = "HTML"  # HTML or Text
    to_recipients: list[str] = []
    cc_recipients: list[str] = []
    bcc_recipients: list[str] = []
    importance: str = "normal"


class SendMailRequest(BaseModel):
    subject: str
    body: str
    body_type: str = "HTML"  # HTML or Text
    to_recipients: list[str]  # Email addresses
    cc_recipients: list[str] = []
    bcc_recipients: list[str] = []
    importance: str = "normal"
    save_to_sent_items: bool = True


class DraftReplyRequest(BaseModel):
    comment: str = ""


class ReplyMailRequest(BaseModel):
    comment: str
    reply_all: bool = False


class ForwardMailRequest(BaseModel):
    comment: str
    to_recipients: list[str]


class UpdateMailRequest(BaseModel):
    is_read: Optional[bool] = None
    flag_status: Optional[str] = None  # notFlagged, complete, flagged
    categories: Optional[list[str]] = None
    body: Optional[str] = None  # Body content (HTML or text)
    body_type: Optional[str] = None  # "HTML" or "Text" (default: "HTML")


class MoveMailRequest(BaseModel):
    destination_folder_id: str


class BatchMoveRequest(BaseModel):
    message_ids: list[str]
    destination_folder_id: str


class BatchDeleteRequest(BaseModel):
    message_ids: list[str]


# Calendar schemas
class Calendar(BaseModel):
    id: str
    name: str
    color: Optional[str] = None
    can_edit: bool = Field(True, alias="canEdit")
    can_share: bool = Field(True, alias="canShare")
    is_default: bool = Field(False, alias="isDefault")
    owner: Optional[dict] = None

    class Config:
        populate_by_name = True


class EventDateTime(BaseModel):
    date_time: str = Field(alias="dateTime")
    time_zone: str = Field(alias="timeZone")

    class Config:
        populate_by_name = True


class Attendee(BaseModel):
    email_address: EmailAddress = Field(alias="emailAddress")
    type: str = "required"
    status: Optional[dict] = None

    class Config:
        populate_by_name = True


class CalendarEvent(BaseModel):
    id: str
    subject: Optional[str] = None
    body: Optional[dict] = None
    start: Optional[EventDateTime] = None
    end: Optional[EventDateTime] = None
    location: Optional[dict] = None
    attendees: list[Attendee] = []
    organizer: Optional[dict] = None
    is_all_day: bool = Field(False, alias="isAllDay")
    is_cancelled: bool = Field(False, alias="isCancelled")
    is_online_meeting: bool = Field(False, alias="isOnlineMeeting")
    online_meeting_url: Optional[str] = Field(None, alias="onlineMeetingUrl")
    recurrence: Optional[dict] = None
    response_status: Optional[dict] = Field(None, alias="responseStatus")
    show_as: Optional[str] = Field(None, alias="showAs")
    importance: str = "normal"
    sensitivity: str = "normal"
    web_link: Optional[str] = Field(None, alias="webLink")

    class Config:
        populate_by_name = True


class CreateEventRequest(BaseModel):
    subject: str
    body: Optional[str] = None
    body_type: str = "HTML"
    start_datetime: datetime
    end_datetime: datetime
    time_zone: str = "UTC"
    location: Optional[str] = None
    attendees: list[str] = []  # Email addresses
    is_all_day: bool = False
    is_online_meeting: bool = False
    recurrence: Optional[dict] = None
    reminder_minutes: int = 15
    show_as: str = "busy"  # free, tentative, busy, oof, workingElsewhere
    importance: str = "normal"


class UpdateEventRequest(BaseModel):
    subject: Optional[str] = None
    body: Optional[str] = None
    body_type: str = "HTML"
    start_datetime: Optional[datetime] = None
    end_datetime: Optional[datetime] = None
    time_zone: str = "UTC"
    location: Optional[str] = None
    attendees: Optional[list[str]] = None
    is_all_day: Optional[bool] = None
    is_online_meeting: Optional[bool] = None
    reminder_minutes: Optional[int] = None
    show_as: Optional[str] = None
    importance: Optional[str] = None


class RespondEventRequest(BaseModel):
    comment: Optional[str] = None
    send_response: bool = True


# OneDrive schemas
class Drive(BaseModel):
    id: str
    name: str
    drive_type: str = Field(alias="driveType")
    owner: Optional[dict] = None
    quota: Optional[dict] = None

    class Config:
        populate_by_name = True


class DriveItem(BaseModel):
    id: str
    name: str
    size: Optional[int] = None
    created_datetime: Optional[datetime] = Field(None, alias="createdDateTime")
    last_modified_datetime: Optional[datetime] = Field(None, alias="lastModifiedDateTime")
    web_url: Optional[str] = Field(None, alias="webUrl")
    folder: Optional[dict] = None
    file: Optional[dict] = None
    parent_reference: Optional[dict] = Field(None, alias="parentReference")

    class Config:
        populate_by_name = True

    @property
    def is_folder(self) -> bool:
        return self.folder is not None


class CreateFolderRequest(BaseModel):
    name: str


class RenameItemRequest(BaseModel):
    name: Optional[str] = None
    parent_id: Optional[str] = None  # For moving


# Attachment schemas
class Attachment(BaseModel):
    id: str
    name: str
    size: int = 0
    content_type: str = Field("application/octet-stream", alias="contentType")
    is_inline: bool = Field(False, alias="isInline")

    class Config:
        populate_by_name = True


# Thread schemas
class MailThread(BaseModel):
    conversation_id: str
    subject: Optional[str] = None
    latest_datetime: Optional[datetime] = None
    message_count: int = 0
    messages: list[MailMessage] = []


# Background job schemas
class BackgroundJobStatus(BaseModel):
    id: str
    job_type: str
    status: str
    progress: int
    total: int
    result: Optional[dict] = None
    error: Optional[str] = None
    created_at: datetime
    updated_at: datetime


# API key schemas
class ApiKeyCreate(BaseModel):
    name: str
    tier: Optional[str] = None          # e.g. "openclaw" — pre-fills permissions from TIER_PERMISSIONS
    permissions: list[str] = []         # explicit list; ignored when tier is set


class ApiKeyResponse(BaseModel):
    id: int
    name: str
    tier: Optional[str] = None
    permissions: list[str]
    created_at: datetime
    last_used_at: Optional[datetime] = None
    is_active: bool


class ApiKeyCreated(ApiKeyResponse):
    """Response when a new key is created — includes the raw key (shown only once)."""
    raw_key: str


class ApiKeyUpdate(BaseModel):
    name: Optional[str] = None
    permissions: Optional[list[str]] = None
    is_active: Optional[bool] = None


# Contact schemas
class Contact(BaseModel):
    id: str
    display_name: Optional[str] = Field(None, alias="displayName")
    given_name: Optional[str] = Field(None, alias="givenName")
    surname: Optional[str] = None
    email_addresses: list[dict] = Field(default_factory=list, alias="emailAddresses")
    mobile_phone: Optional[str] = Field(None, alias="mobilePhone")
    company_name: Optional[str] = Field(None, alias="companyName")
    job_title: Optional[str] = Field(None, alias="jobTitle")
    personal_notes: Optional[str] = Field(None, alias="personalNotes")

    class Config:
        populate_by_name = True


class CreateContactRequest(BaseModel):
    name: str
    email: Optional[str] = None
    phone: Optional[str] = None
    organization: Optional[str] = None
    title: Optional[str] = None
    notes: Optional[str] = None


class UpdateContactRequest(BaseModel):
    name: Optional[str] = None
    email: Optional[str] = None
    phone: Optional[str] = None
    organization: Optional[str] = None
    title: Optional[str] = None
    notes: Optional[str] = None


# Pagination
class PaginatedResponse(BaseModel):
    items: list
    next_link: Optional[str] = None
    total_count: Optional[int] = None
