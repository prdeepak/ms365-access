from fastapi import APIRouter, Depends, HTTPException, BackgroundTasks, Query, Response
from sqlalchemy.ext.asyncio import AsyncSession
from typing import Optional
import logging

from app.database import get_db
from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.mail_service import MailService
from app.schemas import (
    CreateDraftRequest,
    DraftReplyRequest,
    SendMailRequest,
    ReplyMailRequest,
    ForwardMailRequest,
    UpdateMailRequest,
    MoveMailRequest,
    BatchMoveRequest,
    BatchDeleteRequest,
    BackgroundJobStatus,
)
from app.models import Auth, BackgroundJob
from app.tasks.background import create_job, run_batch_operation
from app import audit

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/mail", tags=["mail"])


def get_mail_service(graph_client: GraphClient = Depends(get_graph_client)) -> MailService:
    return MailService(graph_client)


@router.get("/folders", dependencies=[Depends(require_permission("read:mail"))])
async def list_folders(mail_service: MailService = Depends(get_mail_service)):
    result = await mail_service.list_folders()
    return result.get("value", [])


@router.get("/folders/resolve/{name}", dependencies=[Depends(require_permission("read:mail"))])
async def resolve_folder_name(
    name: str, mail_service: MailService = Depends(get_mail_service)
):
    """Resolve a well-known folder name to its folder object including ID.

    Well-known folder names include: inbox, drafts, sentitems, deleteditems,
    junkemail, archive, outbox, etc.

    This is useful because MS Graph API can cache results when querying by
    well-known name. Using the actual folder ID ensures fresh results.
    """
    try:
        folder = await mail_service.resolve_folder_name(name)
        return folder
    except Exception as e:
        raise HTTPException(
            status_code=404,
            detail=f"Could not resolve folder name '{name}': {str(e)}",
        )


@router.get("/messages", dependencies=[Depends(require_permission("read:mail"))])
async def list_messages(
    folder: Optional[str] = Query(
        None,
        description="Well-known folder name (inbox, archive, junkemail, etc.). "
        "Will be resolved to folder_id internally to avoid MS Graph caching issues.",
    ),
    folder_id: Optional[str] = Query(
        None,
        description="Actual folder ID. Takes precedence over 'folder' if both provided.",
    ),
    top: int = Query(25, ge=1),
    skip: int = Query(0, ge=0),
    search: Optional[str] = None,
    filter: Optional[str] = None,
    order_by: str = "receivedDateTime desc",
    include_body: bool = Query(False, description="Include truncated plain-text body (~2000 chars)"),
    mail_service: MailService = Depends(get_mail_service),
):
    """List messages from a folder.

    Note: MS Graph API can cache results when querying folders by well-known name.
    This endpoint resolves folder names to IDs internally to ensure fresh results.
    """
    # MS Graph API caps $top at 100 per page
    top = min(top, 100)

    # Resolve folder name to ID if folder is provided but folder_id is not
    resolved_folder_id = folder_id
    if folder and not folder_id:
        try:
            folder_data = await mail_service.resolve_folder_name(folder)
            resolved_folder_id = folder_data.get("id")
            logger.debug(f"Resolved folder '{folder}' to ID '{resolved_folder_id}'")
        except Exception as e:
            logger.warning(f"Could not resolve folder '{folder}': {e}")
            # Fall back to using the folder name directly
            resolved_folder_id = folder

    result = await mail_service.list_messages(
        folder_id=resolved_folder_id,
        top=top,
        skip=skip,
        search=search,
        filter_query=filter,
        order_by=order_by,
        include_body=include_body,
    )
    return result.get("value", [])


@router.get("/messages/{message_id}", dependencies=[Depends(require_permission("read:mail"))])
async def get_message(
    message_id: str,
    mail_service: MailService = Depends(get_mail_service),
):
    return await mail_service.get_message(message_id)


@router.post("/drafts", dependencies=[Depends(require_permission("write:draft"))], status_code=201)
async def create_draft(
    request: CreateDraftRequest,
    mail_service: MailService = Depends(get_mail_service),
):
    """Create a new draft message in Drafts folder (does not send).

    Requires write:draft permission. To send the draft later, use
    POST /mail/messages/{draft_id}/send (requires write:mail).
    """
    return await mail_service.create_draft(
        subject=request.subject,
        body=request.body,
        body_type=request.body_type,
        to_recipients=request.to_recipients or None,
        cc_recipients=request.cc_recipients or None,
        bcc_recipients=request.bcc_recipients or None,
        importance=request.importance,
    )


@router.post("/messages", dependencies=[Depends(require_permission("write:mail"))])
async def send_mail(
    request: SendMailRequest,
    mail_service: MailService = Depends(get_mail_service),
    auth: Auth = Depends(get_current_auth),
):
    await mail_service.send_mail(
        subject=request.subject,
        body=request.body,
        body_type=request.body_type,
        to_recipients=request.to_recipients,
        cc_recipients=request.cc_recipients,
        bcc_recipients=request.bcc_recipients,
        importance=request.importance,
        save_to_sent_items=request.save_to_sent_items,
    )
    audit.log_mail_send(auth.email, request.to_recipients, request.subject)
    return {"message": "Email sent successfully"}


@router.post("/messages/{message_id}/send", dependencies=[Depends(require_permission("write:mail"))])
async def send_draft(
    message_id: str,
    mail_service: MailService = Depends(get_mail_service),
    auth: Auth = Depends(get_current_auth),
):
    """Send an existing draft message."""
    await mail_service.send_draft(message_id)
    audit.log_mail_send(auth.email, [], f"draft:{message_id}")
    return {"message": "Draft sent successfully"}


@router.post("/messages/{message_id}/draftReply", dependencies=[Depends(require_permission("write:draft"))])
async def create_reply_draft(
    message_id: str,
    reply_all: bool = Query(False, description="Create reply-all draft instead of reply"),
    request: Optional[DraftReplyRequest] = None,
    mail_service: MailService = Depends(get_mail_service),
):
    """Create a draft reply to a message (does not send it).

    Returns the draft message object with its ID. Use PATCH /mail/messages/{draft_id}
    to update the body before sending.
    """
    comment = request.comment if request else ""
    result = await mail_service.create_reply_draft(
        message_id=message_id,
        reply_all=reply_all,
        comment=comment,
    )
    return result


@router.post("/messages/{message_id}/reply", dependencies=[Depends(require_permission("write:mail"))])
async def reply_to_message(
    message_id: str,
    request: ReplyMailRequest,
    mail_service: MailService = Depends(get_mail_service),
):
    await mail_service.reply_to_message(
        message_id=message_id,
        comment=request.comment,
        reply_all=request.reply_all,
    )
    return {"message": "Reply sent successfully"}


@router.post("/messages/{message_id}/forward", dependencies=[Depends(require_permission("write:mail"))])
async def forward_message(
    message_id: str,
    request: ForwardMailRequest,
    mail_service: MailService = Depends(get_mail_service),
):
    await mail_service.forward_message(
        message_id=message_id,
        comment=request.comment,
        to_recipients=request.to_recipients,
    )
    return {"message": "Message forwarded successfully"}


@router.patch("/messages/{message_id}", dependencies=[Depends(require_permission("write:draft"))])
async def update_message(
    message_id: str,
    request: UpdateMailRequest,
    mail_service: MailService = Depends(get_mail_service),
):
    return await mail_service.update_message(
        message_id=message_id,
        is_read=request.is_read,
        flag_status=request.flag_status,
        categories=request.categories,
        body=request.body,
        body_type=request.body_type,
    )


@router.post("/messages/{message_id}/move", dependencies=[Depends(require_permission("write:mail"))])
async def move_message(
    message_id: str,
    request: MoveMailRequest,
    verify: bool = Query(
        True, description="Verify the move by checking parentFolderId after"
    ),
    mail_service: MailService = Depends(get_mail_service),
    auth: Auth = Depends(get_current_auth),
):
    """Move a message to a destination folder.

    The destination_folder_id can be either an actual folder ID or a well-known
    folder name (inbox, archive, junkemail, etc.).

    If verify=True (default), the response will include a 'verified' field
    indicating whether the message's parentFolderId matches the destination.
    """
    result = await mail_service.move_message(
        message_id=message_id,
        destination_folder_id=request.destination_folder_id,
        verify=verify,
    )

    if verify and not result.get("verified", True):
        logger.warning(
            f"Move verification failed for message {message_id} "
            f"to folder {request.destination_folder_id}"
        )

    audit.log_mail_move(auth.email, message_id, request.destination_folder_id)
    return result


@router.delete("/messages/{message_id}", dependencies=[Depends(require_permission("write:mail"))])
async def delete_message(
    message_id: str,
    mail_service: MailService = Depends(get_mail_service),
    auth: Auth = Depends(get_current_auth),
):
    await mail_service.delete_message(message_id)
    audit.log_mail_delete(auth.email, message_id)
    return {"message": "Message deleted successfully"}


@router.post("/batch/move", dependencies=[Depends(require_permission("write:mail"))])
async def batch_move_messages(
    request: BatchMoveRequest,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
    mail_service: MailService = Depends(get_mail_service),
):
    job = await create_job(db, "batch_move", total=len(request.message_ids))

    async def move_operation(message_id: str):
        await mail_service.move_message(message_id, request.destination_folder_id)

    background_tasks.add_task(
        run_batch_operation,
        job.id,
        request.message_ids,
        move_operation,
    )

    return {"job_id": job.id, "message": "Batch move started"}


@router.post("/batch/delete", dependencies=[Depends(require_permission("write:mail"))])
async def batch_delete_messages(
    request: BatchDeleteRequest,
    background_tasks: BackgroundTasks,
    db: AsyncSession = Depends(get_db),
    mail_service: MailService = Depends(get_mail_service),
    auth: Auth = Depends(get_current_auth),
):
    job = await create_job(db, "batch_delete", total=len(request.message_ids))

    async def delete_operation(message_id: str):
        await mail_service.delete_message(message_id)

    background_tasks.add_task(
        run_batch_operation,
        job.id,
        request.message_ids,
        delete_operation,
    )

    audit.log_mail_batch_delete(auth.email, request.message_ids)
    return {"job_id": job.id, "message": "Batch delete started"}


@router.get("/search", dependencies=[Depends(require_permission("read:mail"))])
async def search_messages(
    q: str,
    top: int = Query(25, ge=1),
    mail_service: MailService = Depends(get_mail_service),
):
    top = min(top, 100)
    result = await mail_service.search_messages(query=q, top=top)
    return result.get("value", [])


@router.get("/threads", dependencies=[Depends(require_permission("read:mail"))])
async def list_threads(
    folder: Optional[str] = Query(
        None,
        description="Well-known folder name (inbox, archive, junkemail, etc.).",
    ),
    folder_id: Optional[str] = Query(
        None,
        description="Actual folder ID. Takes precedence over 'folder' if both provided.",
    ),
    top: int = Query(25, ge=1),
    mail_service: MailService = Depends(get_mail_service),
):
    """List messages grouped by conversationId, ordered by most recent thread activity.

    Returns threads with: conversationId, subject, messages[], latestDateTime, messageCount.
    """
    top = min(top, 100)

    # Resolve folder name to ID if needed
    resolved_folder_id = folder_id
    if folder and not folder_id:
        try:
            folder_data = await mail_service.resolve_folder_name(folder)
            resolved_folder_id = folder_data.get("id")
        except Exception as e:
            logger.warning(f"Could not resolve folder '{folder}': {e}")
            resolved_folder_id = folder

    threads = await mail_service.list_threads(
        folder_id=resolved_folder_id,
        top=top,
    )
    return threads


@router.get("/messages/{message_id}/attachments", dependencies=[Depends(require_permission("read:mail"))])
async def list_attachments(
    message_id: str,
    mail_service: MailService = Depends(get_mail_service),
):
    """List attachments for a message (id, name, size, contentType)."""
    attachments = await mail_service.list_attachments(message_id)
    return attachments


@router.get("/messages/{message_id}/attachments/{attachment_id}", dependencies=[Depends(require_permission("read:mail"))])
async def download_attachment(
    message_id: str,
    attachment_id: str,
    mail_service: MailService = Depends(get_mail_service),
):
    """Download attachment content."""
    content = await mail_service.get_attachment(message_id, attachment_id)
    return Response(content=content, media_type="application/octet-stream")
