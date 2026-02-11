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

    def __init__(self, base_url="http://localhost:8365"):
        self.base_url = base_url.rstrip("/")

    # ------------------------------------------------------------------
    # Low-level helpers
    # ------------------------------------------------------------------

    def _get_json(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            with urlopen(Request(url), timeout=timeout) as resp:
                return json.loads(resp.read().decode())
        except (URLError, HTTPError, json.JSONDecodeError) as e:
            log.warning(f"GET {path} failed: {e}")
            return None

    def _get_raw(self, path, params=None, timeout=30):
        url = f"{self.base_url}{path}"
        if params:
            url += ("&" if "?" in path else "?") + urlencode(params)
        try:
            with urlopen(Request(url), timeout=timeout) as resp:
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

    def batch_delete_messages(self, message_ids):
        """Batch Delete Messages"""
        data = {}
        data["message_ids"] = message_ids
        return self._post_json(f"/mail/batch/delete", data)

    def batch_move_messages(self, message_ids, destination_folder_id):
        """Batch Move Messages"""
        data = {}
        data["message_ids"] = message_ids
        data["destination_folder_id"] = destination_folder_id
        return self._post_json(f"/mail/batch/move", data)

    def list_folders(self):
        """List Folders"""
        return self._get_json(f"/mail/folders")

    def resolve_folder_name(self, name):
        """Resolve Folder Name"""
        return self._get_json(f"/mail/folders/resolve/{name}")

    def list_messages(
            self,
            folder=None,
            folder_id=None,
            top=25,
            skip=0,
            search=None,
            filter=None,
            order_by='receivedDateTime desc',
    ):
        """List Messages"""
        params = {k: v for k, v in {"folder": folder, "folder_id": folder_id, "top": top, "skip": skip, "search": search, "filter": filter, "order_by": order_by}.items() if v is not None}
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
        return self._post_json(f"/mail/messages", data)

    def delete_message(self, message_id):
        """Delete Message"""
        return self._delete_json(f"/mail/messages/{message_id}")

    def get_message(self, message_id):
        """Get Message"""
        return self._get_json(f"/mail/messages/{message_id}")

    def update_message(
            self,
            message_id,
            is_read=None,
            flag_status=None,
            categories=None,
            body=None,
            body_type=None,
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
        return self._patch_json(f"/mail/messages/{message_id}", data)

    def create_reply_draft(self, message_id, reply_all=False):
        """Create Reply Draft"""
        params = {k: v for k, v in {"reply_all": reply_all}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/draftReply", params=params)

    def forward_message(self, message_id, comment, to_recipients):
        """Forward Message"""
        data = {}
        data["comment"] = comment
        data["to_recipients"] = to_recipients
        return self._post_json(f"/mail/messages/{message_id}/forward", data)

    def move_message(self, message_id, destination_folder_id, verify=True):
        """Move Message"""
        data = {}
        data["destination_folder_id"] = destination_folder_id
        params = {k: v for k, v in {"verify": verify}.items() if v is not None}
        return self._post_json(f"/mail/messages/{message_id}/move", data, params)

    def reply_to_message(self, message_id, comment, reply_all=False):
        """Reply To Message"""
        data = {}
        data["comment"] = comment
        if reply_all is not None:
            data["reply_all"] = reply_all
        return self._post_json(f"/mail/messages/{message_id}/reply", data)

    def send_draft(self, message_id):
        """Send Draft"""
        return self._post_json(f"/mail/messages/{message_id}/send")

    def search_messages(self, q, top=25, skip=0):
        """Search Messages"""
        params = {k: v for k, v in {"q": q, "top": top, "skip": skip}.items() if v is not None}
        return self._get_json(f"/mail/search", params)
    # ------------------------------------------------------------------
    # Calendar
    # ------------------------------------------------------------------

    def list_calendars(self):
        """List Calendars"""
        return self._get_json(f"/calendar/calendars")

    def list_events(
            self,
            calendar_id=None,
            top=25,
            skip=0,
            order_by='start/dateTime',
            filter=None,
    ):
        """List Events"""
        params = {k: v for k, v in {"calendar_id": calendar_id, "top": top, "skip": skip, "order_by": order_by, "filter": filter}.items() if v is not None}
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

    def get_event(self, event_id):
        """Get Event"""
        return self._get_json(f"/calendar/events/{event_id}")

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
    ):
        """Get Calendar View"""
        params = {k: v for k, v in {"start_datetime": start_datetime, "end_datetime": end_datetime, "calendar_id": calendar_id, "top": top}.items() if v is not None}
        return self._get_json(f"/calendar/view", params)
    # ------------------------------------------------------------------
    # Files
    # ------------------------------------------------------------------

    def get_drive_root(self, drive_id=None):
        """Get Drive Root"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._get_json(f"/files/drive/root", params)

    def files_list_drives(self):
        """List Drives"""
        return self._get_json(f"/files/drives")

    def delete_item(self, item_id, drive_id=None):
        """Delete Item"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._delete_json(f"/files/items/{item_id}", params)

    def files_get_item(self, item_id, drive_id=None):
        """Get Item"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
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
    ):
        """List Children"""
        params = {k: v for k, v in {"drive_id": drive_id, "top": top, "skip": skip, "order_by": order_by}.items() if v is not None}
        return self._get_json(f"/files/items/{item_id}/children", params)

    def files_download_content(self, item_id, drive_id=None):
        """Download Content"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._get_raw(f"/files/items/{item_id}/content", params)

    def create_folder(self, parent_id, name, drive_id=None):
        """Create Folder"""
        data = {}
        data["name"] = name
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._post_json(f"/files/items/{parent_id}/folder", data, params)

    def upload_content(self, parent_id, filename, data=None, drive_id=None):
        """Upload Content"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._put_json(f"/files/items/{parent_id}:/{filename}:/content", data, params)

    def search_files(self, q, drive_id=None, top=25):
        """Search Files"""
        params = {k: v for k, v in {"q": q, "drive_id": drive_id, "top": top}.items() if v is not None}
        return self._get_json(f"/files/search", params)
    # ------------------------------------------------------------------
    # Sharepoint
    # ------------------------------------------------------------------

    def sharepoint_list_drives(self, site_id):
        """List Drives"""
        params = {k: v for k, v in {"site_id": site_id}.items() if v is not None}
        return self._get_json(f"/sharepoint/drives", params)

    def sharepoint_get_item(self, item_id, drive_id):
        """Get Item"""
        params = {k: v for k, v in {"drive_id": drive_id}.items() if v is not None}
        return self._get_json(f"/sharepoint/items/{item_id}", params)

    def sharepoint_list_children(self, item_id, drive_id, top=100, order_by='name'):
        """List Children"""
        params = {k: v for k, v in {"drive_id": drive_id, "top": top, "order_by": order_by}.items() if v is not None}
        return self._get_json(f"/sharepoint/items/{item_id}/children", params)

    def sharepoint_download_content(self, item_id, drive_id, format=None):
        """Download Content"""
        params = {k: v for k, v in {"drive_id": drive_id, "format": format}.items() if v is not None}
        return self._get_raw(f"/sharepoint/items/{item_id}/content", params)

    def resolve_url(self, url):
        """Resolve Url"""
        params = {k: v for k, v in {"url": url}.items() if v is not None}
        return self._get_json(f"/sharepoint/resolve", params)

    def search(self, q, drive_id, top=25):
        """Search"""
        params = {k: v for k, v in {"q": q, "drive_id": drive_id, "top": top}.items() if v is not None}
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
    # Auth
    # ------------------------------------------------------------------

    def logout(self):
        """Logout"""
        return self._post_json(f"/auth/logout")

    def auth_status(self):
        """Auth Status"""
        return self._get_json(f"/auth/status")
