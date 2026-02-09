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
            params["$select"] = "id,subject,bodyPreview,from,toRecipients,ccRecipients,receivedDateTime,sentDateTime,isRead,isDraft,hasAttachments,importance,flag"

        if search:
            params["$search"] = f'"{search}"'

        if filter_query:
            params["$filter"] = filter_query

        return await self.client.get(endpoint, params=params)

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

    async def create_reply_draft(
        self, message_id: str, reply_all: bool = False
    ) -> dict:
        """Create a draft reply to a message (does not send it).

        Uses MS Graph createReply/createReplyAll which creates a draft
        message in the Drafts folder with the proper reply headers.
        """
        endpoint = f"/me/messages/{message_id}/{'createReplyAll' if reply_all else 'createReply'}"
        return await self.client.post(endpoint, {})

    async def search_messages(
        self,
        query: str,
        top: int = 25,
        skip: int = 0,
    ) -> dict:
        params = {
            "$search": f'"{query}"',
            "$top": top,
            "$skip": skip,
            "$select": "id,subject,bodyPreview,from,toRecipients,receivedDateTime,isRead,hasAttachments",
        }
        return await self.client.get("/me/messages", params=params)
