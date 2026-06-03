"""Excel Workbook API service.

Cell/range/table-level writes to an Excel (.xlsx) file via the Microsoft
Graph Excel API. Unlike whole-file replace (see SharePointService /
OneDriveService), these edits go through the same Excel backend that powers
Excel Online co-authoring, so they merge safely into a workbook that is open
on multiple machines.

Reference: https://learn.microsoft.com/graph/api/resources/excel

Notes / constraints (per MS best-practices):
  - .xlsx only, stored in OneDrive/SharePoint. Legacy .xls and very large
    workbooks may fail with GenericFileOpenError.
  - Writes to a single workbook should be SEQUENTIAL. Concurrent writes cause
    merge conflicts and throttling.
  - A persistent session (persistChanges=True) writes through to the file and
    merges with the live co-authoring session.
  - An exclusive check-out by another user blocks writes; we guard for this.
"""

import logging
from typing import Optional
from urllib.parse import quote

import httpx

from app.services.graph_client import GraphClient

logger = logging.getLogger(__name__)


class WorkbookLockedError(Exception):
    """Raised when a workbook can't be written because it is checked out / locked."""


class WorkbookService:
    def __init__(self, graph_client: GraphClient):
        self.client = graph_client

    def _base(self, item_id: str, site_id: Optional[str] = None) -> str:
        """Build the workbook endpoint prefix.

        Uses /sites/{site_id}/drive for SharePoint items, /me/drive for the
        signed-in user's OneDrive when no site is given.
        """
        drive = f"/sites/{site_id}/drive" if site_id else "/me/drive"
        return f"{drive}/items/{item_id}/workbook"

    def _item_path(self, item_id: str, site_id: Optional[str] = None) -> str:
        drive = f"/sites/{site_id}/drive" if site_id else "/me/drive"
        return f"{drive}/items/{item_id}"

    @staticmethod
    def _session_header(session_id: Optional[str]) -> Optional[dict]:
        return {"workbook-session-id": session_id} if session_id else None

    @staticmethod
    def _sheet_ref(sheet: str) -> str:
        """Encode a worksheet name for use in a workbook function-path segment.

        Single quotes in the name must be doubled, then URL-encoded.
        """
        return quote(sheet.replace("'", "''"), safe="")

    # -- Lock / safety preflight ------------------------------------------

    async def get_lock_state(
        self, item_id: str, site_id: Optional[str] = None
    ) -> dict:
        """Best-effort check of whether the file is checked out by someone.

        Co-authoring (multiple people editing) does NOT check a file out, so a
        co-authored file reports checked_out=False and is safe to write. An
        exclusive check-out reports the user holding it.
        """
        item = await self.client.get(
            self._item_path(item_id, site_id),
            params={"$expand": "listItem($expand=fields)"},
        )
        checked_out = False
        checkout_user = None
        list_item = item.get("listItem") or {}
        fields = list_item.get("fields") or {}
        # SharePoint document libraries surface check-out via these columns.
        for key in ("CheckoutUser", "CheckoutUserLookupId"):
            val = fields.get(key)
            if val:
                checked_out = True
                checkout_user = val
        last_modified_by = (
            (item.get("lastModifiedBy") or {}).get("user") or {}
        ).get("displayName")
        return {
            "name": item.get("name"),
            "checked_out": checked_out,
            "checkout_user": checkout_user,
            "last_modified": item.get("lastModifiedDateTime"),
            "last_modified_by": last_modified_by,
            "web_url": item.get("webUrl"),
        }

    async def _ensure_writable(
        self, item_id: str, site_id: Optional[str] = None
    ) -> None:
        """Raise WorkbookLockedError if the file is exclusively checked out."""
        try:
            state = await self.get_lock_state(item_id, site_id)
        except Exception as e:  # best-effort; don't block on a flaky preflight
            logger.warning("Lock preflight failed for %s: %s", item_id, e)
            return
        if state.get("checked_out"):
            user = state.get("checkout_user")
            raise WorkbookLockedError(
                f"'{state.get('name')}' is checked out"
                + (f" (by user ref {user})" if user else "")
                + " — writes are blocked until it is checked in."
            )

    # -- Sessions ----------------------------------------------------------

    async def create_session(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        persist: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Create a workbook session and return its id.

        A persistent session merges writes into the live co-authoring session.
        Pass the returned id to subsequent range/table calls for performance.
        """
        if check_lock:
            await self._ensure_writable(item_id, site_id)
        try:
            return await self.client.post(
                f"{self._base(item_id, site_id)}/createSession",
                data={"persistChanges": persist},
            )
        except httpx.HTTPStatusError as e:
            self._raise_if_locked(e)
            raise

    async def close_session(
        self, item_id: str, session_id: str, site_id: Optional[str] = None
    ) -> dict:
        """Close a workbook session."""
        return await self.client.post(
            f"{self._base(item_id, site_id)}/closeSession",
            data={},
            extra_headers=self._session_header(session_id),
        )

    # -- Discovery ---------------------------------------------------------

    async def list_worksheets(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
    ) -> dict:
        """List worksheets in the workbook."""
        return await self.client.get(
            f"{self._base(item_id, site_id)}/worksheets",
            extra_headers=self._session_header(session_id),
        )

    async def list_tables(
        self,
        item_id: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
    ) -> dict:
        """List tables in the workbook."""
        return await self.client.get(
            f"{self._base(item_id, site_id)}/tables",
            extra_headers=self._session_header(session_id),
        )

    # -- Range read / write ------------------------------------------------

    async def get_range(
        self,
        item_id: str,
        sheet: str,
        address: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
    ) -> dict:
        """Read a cell range, e.g. address='A1:C10'."""
        addr = quote(address, safe=":!$")
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/range(address='{addr}')"
        )
        return await self.client.get(
            endpoint, extra_headers=self._session_header(session_id)
        )

    async def _auto_session(
        self,
        item_id: str,
        site_id: Optional[str],
        session_id: Optional[str],
        auto_session: bool,
        check_lock: bool,
    ):
        """Resolve the session to use for a write.

        Returns (effective_session_id, own_session). When the caller passes no
        session and auto_session is on, open a short-lived persistent session
        so a one-shot write is a clean create -> write -> close in/out. The
        caller must close it (in a finally) when own_session is True.
        """
        if session_id is None and auto_session:
            session = await self.create_session(
                item_id, site_id, persist=True, check_lock=check_lock
            )
            return session.get("id"), True
        if session_id is None and check_lock:
            await self._ensure_writable(item_id, site_id)
        return session_id, False

    async def _close_own(self, item_id, site_id, session_id, own_session) -> None:
        if own_session and session_id:
            try:
                await self.close_session(item_id, session_id, site_id)
            except Exception as e:  # best-effort; the write already landed
                logger.warning("Failed to close auto-session %s: %s", session_id, e)

    async def _session_write(
        self,
        item_id: str,
        site_id: Optional[str],
        session_id: Optional[str],
        auto_session: bool,
        check_lock: bool,
        action,
    ):
        """Run a write `action(effective_session_id)` inside an (optional) session.

        `action` is an async callable taking the effective session id and
        returning the Graph response. Centralises the create -> write -> close
        in/out used by every co-authoring-safe write, with lock-error
        translation. When the caller passes no session and `auto_session` is on,
        a short-lived persistent session is opened and closed around the write.
        """
        session_id, own = await self._auto_session(
            item_id, site_id, session_id, auto_session, check_lock
        )
        try:
            return await action(session_id)
        except httpx.HTTPStatusError as e:
            self._raise_if_locked(e)
            raise
        finally:
            await self._close_own(item_id, site_id, session_id, own)

    async def update_range(
        self,
        item_id: str,
        sheet: str,
        address: str,
        values: list,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Write a 2D array of values into a range.

        `values` must be a list of rows matching the address dimensions, e.g.
        address='A1:B2' -> [["x", 1], ["y", 2]]. Formulas are accepted as
        strings beginning with '='.

        If no `session_id` is given and `auto_session` is True (default), a
        short-lived persistent session is opened and closed around the write —
        the caller just makes one call.
        """
        addr = quote(address, safe=":!$")
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/range(address='{addr}')"
        )

        async def action(sid):
            return await self.client.patch(
                endpoint,
                data={"values": values},
                extra_headers=self._session_header(sid),
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def add_table_row(
        self,
        item_id: str,
        table: str,
        values: list,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Append one or more rows to a table.

        `table` is the table name or id; `values` is a list of rows, e.g.
        [["east", "pear", 4]]. Auto-session behaves as in `update_range`.
        """
        endpoint = (
            f"{self._base(item_id, site_id)}/tables/{quote(table, safe='')}/rows/add"
        )

        async def action(sid):
            return await self.client.post(
                endpoint,
                data={"values": values},
                extra_headers=self._session_header(sid),
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    # -- Worksheet management ----------------------------------------------

    async def get_worksheet(
        self,
        item_id: str,
        sheet: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
    ) -> dict:
        """Read a single worksheet's metadata (id, name, position, visibility)."""
        endpoint = (
            f"{self._base(item_id, site_id)}/worksheets('{self._sheet_ref(sheet)}')"
        )
        return await self.client.get(
            endpoint, extra_headers=self._session_header(session_id)
        )

    async def get_used_range(
        self,
        item_id: str,
        sheet: str,
        values_only: bool = False,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
    ) -> dict:
        """Read the worksheet's used range (the smallest range covering all data).

        With `values_only=True`, formatting-only cells are excluded so the
        range is bounded by cells that actually hold values.
        """
        seg = "usedRange(valuesOnly=true)" if values_only else "usedRange"
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/{seg}"
        )
        return await self.client.get(
            endpoint, extra_headers=self._session_header(session_id)
        )

    async def add_worksheet(
        self,
        item_id: str,
        name: Optional[str] = None,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Add a new worksheet. Graph picks a default name if `name` is omitted."""
        endpoint = f"{self._base(item_id, site_id)}/worksheets/add"
        body = {"name": name} if name else {}

        async def action(sid):
            return await self.client.post(
                endpoint, data=body, extra_headers=self._session_header(sid)
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def delete_worksheet(
        self,
        item_id: str,
        sheet: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Delete a worksheet by name or id."""
        endpoint = (
            f"{self._base(item_id, site_id)}/worksheets('{self._sheet_ref(sheet)}')"
        )

        async def action(sid):
            await self.client.delete(
                endpoint, extra_headers=self._session_header(sid)
            )
            return {"deleted": True, "sheet": sheet}

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def update_worksheet(
        self,
        item_id: str,
        sheet: str,
        name: Optional[str] = None,
        position: Optional[int] = None,
        visibility: Optional[str] = None,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Update a worksheet's properties in one PATCH.

        Any of `name` (rename), `position` (0-based reorder), or `visibility`
        ('Visible' | 'Hidden' | 'VeryHidden') may be set; omitted properties are
        left unchanged.
        """
        endpoint = (
            f"{self._base(item_id, site_id)}/worksheets('{self._sheet_ref(sheet)}')"
        )
        body: dict = {}
        if name is not None:
            body["name"] = name
        if position is not None:
            body["position"] = position
        if visibility is not None:
            body["visibility"] = visibility

        async def action(sid):
            return await self.client.patch(
                endpoint, data=body, extra_headers=self._session_header(sid)
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def copy_worksheet(
        self,
        item_id: str,
        sheet: str,
        name: Optional[str] = None,
        position_type: Optional[str] = None,
        relative_to: Optional[str] = None,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Copy a worksheet, returning the new sheet.

        `position_type` is one of 'None' | 'Before' | 'After' | 'Beginning' |
        'End'; `relative_to` names the anchor sheet for 'Before'/'After'.
        """
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/copy"
        )
        body: dict = {}
        if name:
            body["name"] = name
        if position_type:
            body["positionType"] = position_type
        if relative_to:
            body["relativeTo"] = relative_to

        async def action(sid):
            return await self.client.post(
                endpoint, data=body, extra_headers=self._session_header(sid)
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def protect_worksheet(
        self,
        item_id: str,
        sheet: str,
        options: Optional[dict] = None,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Protect a worksheet. `options` is an optional WorksheetProtectionOptions
        object (e.g. {"allowFormatCells": true}); omit for default protection."""
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/protection/protect"
        )
        body = {"options": options} if options else {}

        async def action(sid):
            return await self.client.post(
                endpoint, data=body, extra_headers=self._session_header(sid)
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def unprotect_worksheet(
        self,
        item_id: str,
        sheet: str,
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Remove protection from a worksheet."""
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/protection/unprotect"
        )

        async def action(sid):
            return await self.client.post(
                endpoint, data={}, extra_headers=self._session_header(sid)
            )

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    async def clear_range(
        self,
        item_id: str,
        sheet: str,
        address: str,
        apply_to: str = "All",
        site_id: Optional[str] = None,
        session_id: Optional[str] = None,
        auto_session: bool = True,
        check_lock: bool = True,
    ) -> dict:
        """Clear a cell range. `apply_to` is 'All' | 'Formats' | 'Contents'."""
        addr = quote(address, safe=":!$")
        endpoint = (
            f"{self._base(item_id, site_id)}"
            f"/worksheets('{self._sheet_ref(sheet)}')/range(address='{addr}')/clear"
        )

        async def action(sid):
            await self.client.post(
                endpoint,
                data={"applyTo": apply_to},
                extra_headers=self._session_header(sid),
            )
            return {"cleared": True, "sheet": sheet, "address": address, "applyTo": apply_to}

        return await self._session_write(
            item_id, site_id, session_id, auto_session, check_lock, action
        )

    # -- Error mapping -----------------------------------------------------

    @staticmethod
    def _raise_if_locked(e: httpx.HTTPStatusError) -> None:
        """Translate a Graph lock/checkout error into WorkbookLockedError."""
        resp = e.response
        if resp is None:
            return
        if resp.status_code == 423:
            raise WorkbookLockedError(
                "The workbook is locked (open exclusively or checked out)."
            ) from e
        try:
            code = (resp.json().get("error") or {}).get("code", "")
        except Exception:
            code = ""
        lock_codes = {
            "resourceLocked", "notAllowed", "accessDenied",
            "lockMismatch", "resourceModified",
        }
        if code in lock_codes:
            raise WorkbookLockedError(
                f"The workbook can't be written right now (Graph: {code})."
            ) from e
