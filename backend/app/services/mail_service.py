import logging
from typing import Optional
from app.services.graph_client import GraphClient

logger = logging.getLogger(__name__)


class MailService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    async def list_folders(self) -> dict:
        return await self.client.get("/me/mailFolders")

    async def get_folder(self, folder_id: str) -> dict:
        """Get a folder by ID or well-known name (inbox, archive, junkemail, etc.)."""
        return await self.client.get(f"/me/mailFolders/{folder_id}")

    async def resolve_folder_name(self, name: str) -> dict:
        """Resolve a well-known folder name to its full folder object including ID.

        Well-known folder names: inbox, drafts, sentitems, deleteditems,
        junkemail, archive, outbox, etc.
        """
        folder = await self.get_folder(name)
        return folder

    async def list_messages(
        self,
        folder_id: Optional[str] = None,
        top: int = 25,
        skip: int = 0,
        search: Optional[str] = None,
        filter_query: Optional[str] = None,
        order_by: str = "receivedDateTime desc",
        select_fields: Optional[str] = None,
        include_body: bool = False,
    ) -> dict:
        if folder_id:
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        else:
            endpoint = "/me/messages"

        params = {
            "$top": top,
            "$skip": skip,
            "$orderby": order_by,
        }

        if select_fields:
            params["$select"] = select_fields
        else:
            fields = "id,subject,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,flag"
            if include_body:
                fields += ",body"
            params["$select"] = fields

        extra_headers = None
        if include_body:
            extra_headers = {"Prefer": 'outlook.body-content-type="text"'}

        if search:
            params["$search"] = f'"{search}"'
            # Graph API rejects $skip and $orderby when combined with $search
            params.pop("$skip", None)
            params.pop("$orderby", None)

        if filter_query:
            params["$filter"] = filter_query

        result = await self.client.get(endpoint, params=params, extra_headers=extra_headers)

        # Truncate body content to 2000 chars when include_body is enabled
        if include_body and "value" in result:
            for msg in result["value"]:
                body = msg.get("body")
                if body and isinstance(body.get("content"), str):
                    body["content"] = body["content"][:2000]

        return result

    async def get_message(self, message_id: str) -> dict:
        return await self.client.get(f"/me/messages/{message_id}")

    async def send_mail(
        self,
        subject: str,
        body: str,
        body_type: str,
        to_recipients: list[str],
        cc_recipients: list[str] = [],
        bcc_recipients: list[str] = [],
        importance: str = "normal",
        save_to_sent_items: bool = True,
    ) -> dict:
        def format_recipients(emails: list[str]) -> list[dict]:
            return [{"emailAddress": {"address": email}} for email in emails]

        message = {
            "subject": subject,
            "body": {
                "contentType": body_type,
                "content": body,
            },
            "toRecipients": format_recipients(to_recipients),
            "importance": importance,
        }

        if cc_recipients:
            message["ccRecipients"] = format_recipients(cc_recipients)
        if bcc_recipients:
            message["bccRecipients"] = format_recipients(bcc_recipients)

        data = {
            "message": message,
            "saveToSentItems": save_to_sent_items,
        }

        return await self.client.post("/me/sendMail", data)

    async def reply_to_message(
        self,
        message_id: str,
        comment: str,
        reply_all: bool = False,
    ) -> dict:
        endpoint = f"/me/messages/{message_id}/{'replyAll' if reply_all else 'reply'}"
        return await self.client.post(endpoint, {"comment": comment})

    async def forward_message(
        self,
        message_id: str,
        comment: str,
        to_recipients: list[str],
    ) -> dict:
        data = {
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": email}} for email in to_recipients],
        }
        return await self.client.post(f"/me/messages/{message_id}/forward", data)

    async def update_message(
        self,
        message_id: str,
        is_read: Optional[bool] = None,
        flag_status: Optional[str] = None,
        categories: Optional[list[str]] = None,
        body: Optional[str] = None,
        body_type: Optional[str] = None,
    ) -> dict:
        data = {}
        if is_read is not None:
            data["isRead"] = is_read
        if flag_status is not None:
            data["flag"] = {"flagStatus": flag_status}
        if categories is not None:
            data["categories"] = categories
        if body is not None:
            data["body"] = {
                "contentType": body_type or "HTML",
                "content": body,
            }

        return await self.client.patch(f"/me/messages/{message_id}", data)

    async def move_message(
        self, message_id: str, destination_folder_id: str, verify: bool = True
    ) -> dict:
        """Move a message to a destination folder.

        Args:
            message_id: The ID of the message to move
            destination_folder_id: The destination folder ID or well-known name
            verify: If True, verify the move by checking parentFolderId after

        Returns:
            dict with the moved message data and 'verified' key if verify=True
        """
        # Resolve destination folder name to ID if needed (for verification)
        resolved_folder_id = destination_folder_id
        if verify:
            try:
                folder = await self.get_folder(destination_folder_id)
                resolved_folder_id = folder.get("id", destination_folder_id)
            except Exception:
                # If we can't resolve, use the provided ID as-is
                pass

        result = await self.client.post(
            f"/me/messages/{message_id}/move",
            {"destinationId": destination_folder_id},
        )

        # Verify the move if requested
        if verify:
            try:
                # The move endpoint returns the moved message, check its parentFolderId
                actual_parent = result.get("parentFolderId")
                verified = actual_parent == resolved_folder_id
                if not verified:
                    logger.warning(
                        f"Move verification failed for message {message_id}: "
                        f"expected parentFolderId={resolved_folder_id}, got {actual_parent}"
                    )
                result["verified"] = verified
            except Exception as e:
                logger.error(f"Move verification error for message {message_id}: {e}")
                result["verified"] = False

        return result

    async def delete_message(self, message_id: str) -> None:
        await self.client.delete(f"/me/messages/{message_id}")

    async def send_draft(self, message_id: str) -> None:
        """Send an existing draft message."""
        await self.client.post(f"/me/messages/{message_id}/send", {})

    async def create_draft(
        self,
        subject: str,
        body: str = "",
        body_type: str = "HTML",
        to_recipients: list[str] | None = None,
        cc_recipients: list[str] | None = None,
        bcc_recipients: list[str] | None = None,
        importance: str = "normal",
    ) -> dict:
        """Create a new draft message in the Drafts folder (does not send it).

        Uses POST /me/messages, which saves to Drafts without sending.
        """
        def _addr(addresses: list[str]) -> list[dict]:
            return [{"emailAddress": {"address": a}} for a in addresses]

        payload: dict = {
            "subject": subject,
            "body": {"contentType": body_type, "content": body},
            "importance": importance,
        }
        if to_recipients:
            payload["toRecipients"] = _addr(to_recipients)
        if cc_recipients:
            payload["ccRecipients"] = _addr(cc_recipients)
        if bcc_recipients:
            payload["bccRecipients"] = _addr(bcc_recipients)

        return await self.client.post("/me/messages", payload)

    async def create_reply_draft(
        self, message_id: str, reply_all: bool = False, comment: str = ""
    ) -> dict:
        """Create a draft reply to a message (does not send it).

        Uses MS Graph createReply/createReplyAll which creates a draft
        message in the Drafts folder with the proper reply headers.
        The optional comment is inserted as the reply body text.
        """
        endpoint = f"/me/messages/{message_id}/{'createReplyAll' if reply_all else 'createReply'}"
        data = {"comment": comment} if comment else {}
        return await self.client.post(endpoint, data)

    async def search_messages(
        self,
        query: str,
        top: int = 25,
    ) -> dict:
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$select": "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments",
        }
        return await self.client.get("/me/messages", params=params)

    async def list_threads(
        self,
        folder_id: Optional[str] = None,
        top: int = 25,
    ) -> list[dict]:
        """List messages grouped by conversationId, ordered by most recent thread activity.

        Fetches a larger batch of messages (up to 250) and groups them client-side
        by conversationId, then returns the requested number of threads.
        """
        # Fetch more messages than requested threads to get enough conversations
        fetch_count = min(top * 10, 250)

        if folder_id:
            endpoint = f"/me/mailFolders/{folder_id}/messages"
        else:
            endpoint = "/me/messages"

        params = {
            "$top": fetch_count,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,bodyPreview,from,toRecipients,ccRecipients,"
                       "receivedDateTime,sentDateTime,isRead,isDraft,"
                       "hasAttachments,importance,flag,conversationId",
        }

        result = await self.client.get(endpoint, params=params)
        messages = result.get("value", [])

        # Group by conversationId
        threads: dict[str, dict] = {}
        for msg in messages:
            conv_id = msg.get("conversationId")
            if not conv_id:
                continue

            if conv_id not in threads:
                threads[conv_id] = {
                    "conversationId": conv_id,
                    "subject": msg.get("subject"),
                    "latestDateTime": msg.get("receivedDateTime"),
                    "messageCount": 0,
                    "messages": [],
                }

            threads[conv_id]["messages"].append(msg)
            threads[conv_id]["messageCount"] += 1

            # Track the most recent datetime
            msg_dt = msg.get("receivedDateTime")
            if msg_dt and (
                threads[conv_id]["latestDateTime"] is None
                or msg_dt > threads[conv_id]["latestDateTime"]
            ):
                threads[conv_id]["latestDateTime"] = msg_dt

        # Sort threads by most recent activity (descending)
        sorted_threads = sorted(
            threads.values(),
            key=lambda t: t["latestDateTime"] or "",
            reverse=True,
        )

        return sorted_threads[:top]

    async def list_attachments(self, message_id: str) -> list[dict]:
        """List attachments for a message."""
        result = await self.client.get(
            f"/me/messages/{message_id}/attachments",
            params={"$select": "id,name,size,contentType,isInline"},
        )
        return result.get("value", [])

    async def get_attachment(self, message_id: str, attachment_id: str) -> bytes:
        """Download attachment content (raw bytes)."""
        return await self.client.get_raw(
            f"/me/messages/{message_id}/attachments/{attachment_id}/$value"
        )
