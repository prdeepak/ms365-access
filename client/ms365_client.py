"""Python client for ms365-access (http://localhost:8365).

Auto-generated from OpenAPI spec — do not edit manually.
Regenerate with: make gen-client

Zero dependencies beyond stdlib.

Usage:
    from ms365_client import Ms365Client
    client = Ms365Client()  # defaults to http://localhost:8365
"""

import json
import logging
from urllib.error import HTTPError, URLError
from urllib.parse import urlencode
from urllib.request import Request, urlopen

log = logging.getLogger("ms365_client")


class Ms365Client:
    """Client for Ms365 Access HTTP API."""

    def __init__(self, base_url="http://localhost:8365", api_key=None):
        self.base_url = base_url.rstrip("/")
        self.api_key = api_key

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

    def _auth_header(self, req):
        """Add Authorization header if api_key is set."""
        if self.api_key:
            req.add_header("Authorization", f"Bearer {self.api_key}")

    def _get_json(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            req = Request(url)
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"GET {path} failed: {e}")
            return None

    def _get_raw(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            req = Request(url)
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return resp.read()
        except (URLError, HTTPError) as e:
            log.warning(f"GET (raw) {path} failed: {e}")
            return None

    def _post_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="POST")
            req.add_header("Content-Type", "application/json")
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"POST {path} failed: {e}")
            return None

    def _put_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="PUT")
            req.add_header("Content-Type", "application/json")
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"PUT {path} failed: {e}")
            return None

    def _patch_json(self, path, data=None, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            body = json.dumps(data).encode() if data else b""
            req = Request(url, data=body, method="PATCH")
            req.add_header("Content-Type", "application/json")
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"PATCH {path} failed: {e}")
            return None

    def _delete_json(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            req = Request(url, method="DELETE")
            self._auth_header(req)
            with urlopen(req, timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"DELETE {path} failed: {e}")
            return None

    # ------------------------------------------------------------------
    # Health
    # ------------------------------------------------------------------

    def health(self):
        """Health"""
        return self._get_json(f"/health")
    # ------------------------------------------------------------------
    # Mail
    # ------------------------------------------------------------------

    def batch_delete_messages(self, message_ids, user=None):
        """Batch Delete Messages"""
        data = {}
        data["message_ids"] = message_ids
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/batch/delete", data, params)

    def batch_move_messages(self, message_ids, destination_folder_id, user=None):
        """Batch Move Messages"""
        data = {}
        data["message_ids"] = message_ids
        data["destination_folder_id"] = destination_folder_id
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/batch/move", data, params)

    def create_draft(
            self,
            subject,
            body='',
            body_type='HTML',
            to_recipients=None,
            cc_recipients=None,
            bcc_recipients=None,
            importance='normal',
            user=None,
    ):
        """Create Draft"""
        data = {}
        data["subject"] = subject
        if body is not None:
            data["body"] = body
        if body_type is not None:
            data["body_type"] = body_type
        if to_recipients is not None:
            data["to_recipients"] = to_recipients
        if cc_recipients is not None:
            data["cc_recipients"] = cc_recipients
        if bcc_recipients is not None:
            data["bcc_recipients"] = bcc_recipients
        if importance is not None:
            data["importance"] = importance
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/drafts", data, params)

    def list_folders(self, user=None):
        """List Folders"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/mail/folders", params)

    def resolve_folder_name(self, name, user=None):
        """Resolve Folder Name"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/mail/folders/resolve/{name}", params)

    def list_messages(
            self,
            folder=None,
            folder_id=None,
            top=25,
            skip=0,
            search=None,
            filter=None,
            order_by='receivedDateTime desc',
            include_body=False,
            user=None,
    ):
        """List Messages"""
        params = {k: v for k, v in {"folder": folder, "folder_id": folder_id, "top": top, "skip": skip, "search": search, "filter": filter, "order_by": order_by, "include_body": include_body, "user": user}.items() if v is not None}
        return self._get_json(f"/mail/messages", params)

    def send_mail(
            self,
            subject,
            body,
            to_recipients,
            body_type='HTML',
            cc_recipients=None,
            bcc_recipients=None,
            importance='normal',
            save_to_sent_items=True,
            user=None,
    ):
        """Send Mail"""
        data = {}
        data["subject"] = subject
        data["body"] = body
        if body_type is not None:
            data["body_type"] = body_type
        data["to_recipients"] = to_recipients
        if cc_recipients is not None:
            data["cc_recipients"] = cc_recipients
        if bcc_recipients is not None:
            data["bcc_recipients"] = bcc_recipients
        if importance is not None:
            data["importance"] = importance
        if save_to_sent_items is not None:
            data["save_to_sent_items"] = save_to_sent_items
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages", data, params)

    def delete_message(self, message_id, user=None):
        """Delete Message"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._delete_json(f"/mail/messages/{message_id}", params)

    def get_message(self, message_id, user=None):
        """Get Message"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/mail/messages/{message_id}", params)

    def update_message(
            self,
            message_id,
            is_read=None,
            flag_status=None,
            categories=None,
            body=None,
            body_type=None,
            subject=None,
            to_recipients=None,
            cc_recipients=None,
            user=None,
    ):
        """Update Message"""
        data = {}
        if is_read is not None:
            data["is_read"] = is_read
        if flag_status is not None:
            data["flag_status"] = flag_status
        if categories is not None:
            data["categories"] = categories
        if body is not None:
            data["body"] = body
        if body_type is not None:
            data["body_type"] = body_type
        if subject is not None:
            data["subject"] = subject
        if to_recipients is not None:
            data["to_recipients"] = to_recipients
        if cc_recipients is not None:
            data["cc_recipients"] = cc_recipients
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._patch_json(f"/mail/messages/{message_id}", data, params)

    def list_attachments(self, message_id, user=None):
        """List Attachments"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/mail/messages/{message_id}/attachments", params)

    def add_attachment(
            self,
            message_id,
            name,
            content_bytes,
            content_type='application/octet-stream',
            user=None,
    ):
        """Add Attachment"""
        data = {}
        data["name"] = name
        data["content_bytes"] = content_bytes
        if content_type is not None:
            data["content_type"] = content_type
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/attachments", data, params)

    def download_attachment(self, message_id, attachment_id, user=None):
        """Download Attachment"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/mail/messages/{message_id}/attachments/{attachment_id}", params)

    def create_reply_draft(self, message_id, data=None, reply_all=False, user=None):
        """Create Reply Draft"""
        params = {k: v for k, v in {"reply_all": reply_all, "user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/draftReply", data, params)

    def forward_message(self, message_id, comment, to_recipients, user=None):
        """Forward Message"""
        data = {}
        data["comment"] = comment
        data["to_recipients"] = to_recipients
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/forward", data, params)

    def move_message(self, message_id, destination_folder_id, verify=True, user=None):
        """Move Message"""
        data = {}
        data["destination_folder_id"] = destination_folder_id
        params = {k: v for k, v in {"verify": verify, "user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/move", data, params)

    def reply_to_message(self, message_id, comment, reply_all=False, user=None):
        """Reply To Message"""
        data = {}
        data["comment"] = comment
        if reply_all is not None:
            data["reply_all"] = reply_all
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/reply", data, params)

    def send_draft(self, message_id, user=None):
        """Send Draft"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/send", params=params)

    def search_messages(self, q, top=25, user=None):
        """Search Messages"""
        params = {k: v for k, v in {"q": q, "top": top, "user": user}.items() if v is not None}
        return self._get_json(f"/mail/search", params)

    def list_threads(self, folder=None, folder_id=None, top=25, user=None):
        """List Threads"""
        params = {k: v for k, v in {"folder": folder, "folder_id": folder_id, "top": top, "user": user}.items() if v is not None}
        return self._get_json(f"/mail/threads", params)
    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------

    def list_calendars(self, user=None):
        """List Calendars"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/calendar/calendars", params)

    def list_events(
            self,
            calendar_id=None,
            top=25,
            skip=0,
            order_by='start/dateTime',
            filter=None,
            user=None,
    ):
        """List Events"""
        params = {k: v for k, v in {"calendar_id": calendar_id, "top": top, "skip": skip, "order_by": order_by, "filter": filter, "user": user}.items() if v is not None}
        return self._get_json(f"/calendar/events", params)

    def create_event(
            self,
            subject,
            start_datetime,
            end_datetime,
            body=None,
            body_type='HTML',
            time_zone='UTC',
            location=None,
            attendees=None,
            is_all_day=False,
            is_online_meeting=False,
            recurrence=None,
            reminder_minutes=15,
            show_as='busy',
            importance='normal',
            calendar_id=None,
    ):
        """Create Event"""
        data = {}
        data["subject"] = subject
        if body is not None:
            data["body"] = body
        if body_type is not None:
            data["body_type"] = body_type
        data["start_datetime"] = start_datetime
        data["end_datetime"] = end_datetime
        if time_zone is not None:
            data["time_zone"] = time_zone
        if location is not None:
            data["location"] = location
        if attendees is not None:
            data["attendees"] = attendees
        if is_all_day is not None:
            data["is_all_day"] = is_all_day
        if is_online_meeting is not None:
            data["is_online_meeting"] = is_online_meeting
        if recurrence is not None:
            data["recurrence"] = recurrence
        if reminder_minutes is not None:
            data["reminder_minutes"] = reminder_minutes
        if show_as is not None:
            data["show_as"] = show_as
        if importance is not None:
            data["importance"] = importance
        params = {k: v for k, v in {"calendar_id": calendar_id}.items() if v is not None}
        return self._post_json(f"/calendar/events", data, params)

    def delete_event(self, event_id):
        """Delete Event"""
        return self._delete_json(f"/calendar/events/{event_id}")

    def get_event(self, event_id, user=None):
        """Get Event"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/calendar/events/{event_id}", params)

    def update_event(
            self,
            event_id,
            subject=None,
            body=None,
            body_type='HTML',
            start_datetime=None,
            end_datetime=None,
            time_zone='UTC',
            location=None,
            attendees=None,
            is_all_day=None,
            is_online_meeting=None,
            reminder_minutes=None,
            show_as=None,
            importance=None,
    ):
        """Update Event"""
        data = {}
        if subject is not None:
            data["subject"] = subject
        if body is not None:
            data["body"] = body
        if body_type is not None:
            data["body_type"] = body_type
        if start_datetime is not None:
            data["start_datetime"] = start_datetime
        if end_datetime is not None:
            data["end_datetime"] = end_datetime
        if time_zone is not None:
            data["time_zone"] = time_zone
        if location is not None:
            data["location"] = location
        if attendees is not None:
            data["attendees"] = attendees
        if is_all_day is not None:
            data["is_all_day"] = is_all_day
        if is_online_meeting is not None:
            data["is_online_meeting"] = is_online_meeting
        if reminder_minutes is not None:
            data["reminder_minutes"] = reminder_minutes
        if show_as is not None:
            data["show_as"] = show_as
        if importance is not None:
            data["importance"] = importance
        return self._patch_json(f"/calendar/events/{event_id}", data)

    def accept_event(self, event_id, comment=None, send_response=True):
        """Accept Event"""
        data = {}
        if comment is not None:
            data["comment"] = comment
        if send_response is not None:
            data["send_response"] = send_response
        return self._post_json(f"/calendar/events/{event_id}/accept", data)

    def decline_event(self, event_id, comment=None, send_response=True):
        """Decline Event"""
        data = {}
        if comment is not None:
            data["comment"] = comment
        if send_response is not None:
            data["send_response"] = send_response
        return self._post_json(f"/calendar/events/{event_id}/decline", data)

    def tentatively_accept_event(self, event_id, comment=None, send_response=True):
        """Tentatively Accept Event"""
        data = {}
        if comment is not None:
            data["comment"] = comment
        if send_response is not None:
            data["send_response"] = send_response
        return self._post_json(f"/calendar/events/{event_id}/tentative", data)

    def get_calendar_view(
            self,
            start_datetime,
            end_datetime,
            calendar_id=None,
            top=100,
            user=None,
    ):
        """Get Calendar View"""
        params = {k: v for k, v in {"start_datetime": start_datetime, "end_datetime": end_datetime, "calendar_id": calendar_id, "top": top, "user": user}.items() if v is not None}
        return self._get_json(f"/calendar/view", params)
    # ------------------------------------------------------------------
    # Files
    # ------------------------------------------------------------------

    def get_drive_root(self, drive_id=None, user=None):
        """Get Drive Root"""
        params = {k: v for k, v in {"drive_id": drive_id, "user": user}.items() if v is not None}
        return self._get_json(f"/files/drive/root", params)

    def files_list_drives(self, user=None):
        """List Drives"""
        params = {k: v for k, v in {"user": user}.items() if v is not None}
        return self._get_json(f"/files/drives", params)

    def delete_item(self, item_id, drive_id=None):
        """Delete Item"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._delete_json(f"/files/items/{item_id}", params)

    def files_get_item(self, item_id, drive_id=None, user=None):
        """Get Item"""
        params = {k: v for k, v in {"drive_id": drive_id, "user": user}.items() if v is not None}
        return self._get_json(f"/files/items/{item_id}", params)

    def update_item(self, item_id, name=None, parent_id=None, drive_id=None):
        """Update Item"""
        data = {}
        if name is not None:
            data["name"] = name
        if parent_id is not None:
            data["parent_id"] = parent_id
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._patch_json(f"/files/items/{item_id}", data, params)

    def files_list_children(
            self,
            item_id,
            drive_id=None,
            top=100,
            skip=0,
            order_by='name',
            user=None,
    ):
        """List Children"""
        params = {k: v for k, v in {"drive_id": drive_id, "top": top, "skip": skip, "order_by": order_by, "user": user}.items() if v is not None}
        return self._get_json(f"/files/items/{item_id}/children", params)

    def files_download_content(self, item_id, drive_id=None, user=None):
        """Download Content"""
        params = {k: v for k, v in {"drive_id": drive_id, "user": user}.items() if v is not None}
        return self._get_raw(f"/files/items/{item_id}/content", params)

    def files_replace_content(self, item_id, data=None, drive_id=None):
        """Replace Content"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._put_json(f"/files/items/{item_id}/content", data, params)

    def create_folder(self, parent_id, name, drive_id=None):
        """Create Folder"""
        data = {}
        data["name"] = name
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._post_json(f"/files/items/{parent_id}/folder", data, params)

    def files_upload_content(self, parent_id, filename, data=None, drive_id=None):
        """Upload Content"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._put_json(f"/files/items/{parent_id}:/{filename}:/content", data, params)

    def search_files(self, q, drive_id=None, top=25, user=None):
        """Search Files"""
        params = {k: v for k, v in {"q": q, "drive_id": drive_id, "top": top, "user": user}.items() if v is not None}
        return self._get_json(f"/files/search", params)
    # ------------------------------------------------------------------
    # Sharepoint
    # ------------------------------------------------------------------

    def sharepoint_list_drives(self, site_id):
        """List Drives"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._get_json(f"/sharepoint/drives", params)

    def sharepoint_get_item(self, item_id, site_id):
        """Get Item"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._get_json(f"/sharepoint/items/{item_id}", params)

    def rename_item(self, item_id, site_id, data=None):
        """Rename Item"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._patch_json(f"/sharepoint/items/{item_id}", data, params)

    def sharepoint_list_children(self, item_id, site_id, top=100, order_by='name'):
        """List Children"""
        params = {k: v for k, v in {"site_id": site_id, "top": top, "order_by": order_by}.items() if v is not None}
        return self._get_json(f"/sharepoint/items/{item_id}/children", params)

    def sharepoint_download_content(self, item_id, site_id, format=None):
        """Download Content"""
        params = {k: v for k, v in {"site_id": site_id, "format": format}.items() if v is not None}
        return self._get_raw(f"/sharepoint/items/{item_id}/content", params)

    def sharepoint_replace_content(self, item_id, site_id, data=None):
        """Replace Content"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._put_json(f"/sharepoint/items/{item_id}/content", data, params)

    def move_item(self, item_id, site_id, data=None):
        """Move Item"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._patch_json(f"/sharepoint/items/{item_id}/move", data, params)

    def list_versions(self, item_id, site_id, top=50):
        """List Versions"""
        params = {k: v for k, v in {"site_id": site_id, "top": top}.items() if v is not None}
        return self._get_json(f"/sharepoint/items/{item_id}/versions", params)

    def download_version(self, item_id, version_id, site_id):
        """Download Version"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._get_raw(f"/sharepoint/items/{item_id}/versions/{version_id}/content", params)

    def sharepoint_upload_content(self, parent_id, filename, site_id, data=None):
        """Upload Content"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._put_json(f"/sharepoint/items/{parent_id}:/{filename}:/content", data, params)

    def resolve_url(self, url):
        """Resolve Url"""
        params = {k: v for k, v in {"url": url}.items() if v is not None}
        return self._get_json(f"/sharepoint/resolve", params)

    def search(self, q, site_id, top=25):
        """Search"""
        params = {k: v for k, v in {"q": q, "site_id": site_id, "top": top}.items() if v is not None}
        return self._get_json(f"/sharepoint/search", params)

    def resolve_site(self, host_path):
        """Resolve Site"""
        return self._get_json(f"/sharepoint/sites/{host_path}")
    # ------------------------------------------------------------------
    # 
    # ------------------------------------------------------------------

    def root(self):
        """Root"""
        return self._get_json(f"/")
    # ------------------------------------------------------------------
    # Api Keys
    # ------------------------------------------------------------------

    def list_api_keys(self):
        """List Api Keys"""
        return self._get_json(f"/api-keys")

    def create_api_key(self, name, tier=None, permissions=None):
        """Create Api Key"""
        data = {}
        data["name"] = name
        if tier is not None:
            data["tier"] = tier
        if permissions is not None:
            data["permissions"] = permissions
        return self._post_json(f"/api-keys", data)

    def revoke_api_key(self, key_id):
        """Revoke Api Key"""
        return self._delete_json(f"/api-keys/{key_id}")

    def update_api_key(self, key_id, name=None, permissions=None, is_active=None):
        """Update Api Key"""
        data = {}
        if name is not None:
            data["name"] = name
        if permissions is not None:
            data["permissions"] = permissions
        if is_active is not None:
            data["is_active"] = is_active
        return self._patch_json(f"/api-keys/{key_id}", data)
    # ------------------------------------------------------------------
    # Auth
    # ------------------------------------------------------------------

    def logout(self):
        """Logout"""
        return self._post_json(f"/auth/logout")

    def auth_status(self):
        """Auth Status"""
        return self._get_json(f"/auth/status")
    # ------------------------------------------------------------------
    # Contacts
    # ------------------------------------------------------------------

    def list_contacts(self, top=100, skip=0, search=None):
        """List Contacts"""
        params = {k: v for k, v in {"top": top, "skip": skip, "search": search}.items() if v is not None}
        return self._get_json(f"/contacts/", params)

    def create_contact(
            self,
            name,
            email=None,
            phone=None,
            organization=None,
            title=None,
            notes=None,
    ):
        """Create Contact"""
        data = {}
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
        return self._post_json(f"/contacts/", data)

    def search_by_email(self, email):
        """Search By Email"""
        return self._get_json(f"/contacts/by-email/{email}")

    def delete_contact(self, contact_id):
        """Delete Contact"""
        return self._delete_json(f"/contacts/{contact_id}")

    def get_contact(self, contact_id):
        """Get Contact"""
        return self._get_json(f"/contacts/{contact_id}")

    def update_contact(
            self,
            contact_id,
            name=None,
            email=None,
            phone=None,
            organization=None,
            title=None,
            notes=None,
    ):
        """Update Contact"""
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
        return self._patch_json(f"/contacts/{contact_id}", data)
    # ------------------------------------------------------------------
    # Powerbi
    # ------------------------------------------------------------------

    def list_workspaces(self):
        """List Workspaces"""
        return self._get_json(f"/powerbi/workspaces")

    def list_datasets(self, workspace_id):
        """List Datasets"""
        return self._get_json(f"/powerbi/workspaces/{workspace_id}/datasets")

    def execute_query(self, workspace_id, dataset_id, dax_query):
        """Execute Query"""
        data = {}
        data["dax_query"] = dax_query
        return self._post_json(f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/query", data)

    def list_refreshes(self, workspace_id, dataset_id, top=10):
        """List Refreshes"""
        params = {k: v for k, v in {"top": top}.items() if v is not None}
        return self._get_json(f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/refreshes", params)

    def trigger_refresh(self, workspace_id, dataset_id):
        """Trigger Refresh"""
        return self._post_json(f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/refreshes")

    def powerbi_list_tables(self, workspace_id, dataset_id):
        """List Tables"""
        return self._get_json(f"/powerbi/workspaces/{workspace_id}/datasets/{dataset_id}/tables")

    def list_reports(self, workspace_id):
        """List Reports"""
        return self._get_json(f"/powerbi/workspaces/{workspace_id}/reports")
    # ------------------------------------------------------------------
    # Workbook
    # ------------------------------------------------------------------

    def lock_state(self, item_id, site_id=None):
        """Lock State"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._get_json(f"/workbook/items/{item_id}/lock-state", params)

    def get_range(self, item_id, sheet, address, site_id=None, session_id=None):
        """Get Range"""
        params = {k: v for k, v in {"sheet": sheet, "address": address, "site_id": site_id, "session_id": session_id}.items() if v is not None}
        return self._get_json(f"/workbook/items/{item_id}/range", params)

    def update_range(
            self,
            item_id,
            data=None,
            site_id=None,
            session_id=None,
            auto_session=True,
    ):
        """Update Range"""
        params = {k: v for k, v in {"site_id": site_id, "session_id": session_id, "auto_session": auto_session}.items() if v is not None}
        return self._patch_json(f"/workbook/items/{item_id}/range", data, params)

    def close_session(self, item_id, session_id, site_id=None):
        """Close Session"""
        params = {k: v for k, v in {"session_id": session_id, "site_id": site_id}.items() if v is not None}
        return self._delete_json(f"/workbook/items/{item_id}/session", params)

    def create_session(self, item_id, site_id=None, persist=True):
        """Create Session"""
        params = {k: v for k, v in {"site_id": site_id, "persist": persist}.items() if v is not None}
        return self._post_json(f"/workbook/items/{item_id}/session", params=params)

    def workbook_list_tables(self, item_id, site_id=None, session_id=None):
        """List Tables"""
        params = {k: v for k, v in {"site_id": site_id, "session_id": session_id}.items() if v is not None}
        return self._get_json(f"/workbook/items/{item_id}/tables", params)

    def add_table_row(
            self,
            item_id,
            table,
            data=None,
            site_id=None,
            session_id=None,
            auto_session=True,
    ):
        """Add Table Row"""
        params = {k: v for k, v in {"site_id": site_id, "session_id": session_id, "auto_session": auto_session}.items() if v is not None}
        return self._post_json(f"/workbook/items/{item_id}/tables/{table}/rows", data, params)

    def list_worksheets(self, item_id, site_id=None, session_id=None):
        """List Worksheets"""
        params = {k: v for k, v in {"site_id": site_id, "session_id": session_id}.items() if v is not None}
        return self._get_json(f"/workbook/items/{item_id}/worksheets", params)
