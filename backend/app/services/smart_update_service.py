"""Smart workbook updater — the escalation ladder for `files_smart_update`.

Primary path stays whole-file replace-in-place. When the file is open in Excel
(PUT returns 423 Locked), this falls back to surgically updating only the
values/formulas/structure in the live workbook via a Graph Excel session,
leaving formatting alone — or cleanly **defers** (so a caller can notify + retry
on a closed-file replace) when the change can't be reproduced live.

The diff/classify brain lives in :mod:`app.services.workbook_diff` (pure,
Graph-free). This module is the thin orchestrator that drives Graph.

It never notifies and never raises on a lock/complexity wall — it returns
``{"mode": "deferred", "reason": ...}`` and lets the caller decide.
"""

import logging

import httpx

from app.services.onedrive_service import OneDriveService
from app.services.workbook_service import WorkbookService, WorkbookLockedError
from app.services import workbook_diff

logger = logging.getLogger(__name__)

_XLSX_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)


class SmartUpdateService:
    def __init__(
        self,
        onedrive: OneDriveService,
        workbook: WorkbookService,
    ):
        self.onedrive = onedrive
        self.workbook = workbook

    async def smart_update(
        self,
        item_id: str,
        new_bytes: bytes,
        drive_id: str = None,
        site_id: str = None,
        region_map: dict = None,
    ) -> dict:
        """Run the escalation ladder. Returns {mode, ranges_written, reason}.

        mode ∈ {"replaced", "live-edited", "deferred"}.
        """
        # 1. Try full replace-in-place (the common case: nobody has it open).
        try:
            await self.onedrive.replace_content(
                item_id=item_id,
                content=new_bytes,
                content_type=_XLSX_CONTENT_TYPE,
                drive_id=drive_id,
            )
            return {"mode": "replaced", "ranges_written": 0, "reason": "not locked"}
        except httpx.HTTPStatusError as e:
            if e.response is None or e.response.status_code != 423:
                raise  # a real error, not a lock — surface it
            logger.info("smart_update: %s is locked (423); probing live-edit", item_id)

        # 2. Probe live-editability — a session succeeds for an Excel Online
        #    co-authoring lock, fails for an exclusive (Desktop/checked-out) lock.
        try:
            session = await self.workbook.create_session(
                item_id, site_id=site_id, drive_id=drive_id,
                persist=True, check_lock=True,
            )
        except (WorkbookLockedError, httpx.HTTPStatusError):
            return {"mode": "deferred", "ranges_written": 0, "reason": "exclusive lock"}
        session_id = session.get("id")

        try:
            # 3. Download the current live workbook and diff against the proposal.
            live_bytes = await self.onedrive.download_content(
                item_id=item_id, drive_id=drive_id
            )
            plan = workbook_diff.classify(new_bytes, live_bytes, region_map)

            # 4. Defer if the change isn't safely live-editable.
            if plan.mode == "DEFER":
                return {"mode": "deferred", "ranges_written": 0, "reason": plan.reason}

            # 5. Apply the plan in this one session.
            await self._apply(item_id, drive_id, site_id, session_id, plan)
            return {
                "mode": "live-edited",
                "ranges_written": plan.ranges_written,
                "reason": plan.reason,
            }
        except WorkbookLockedError as e:
            return {"mode": "deferred", "ranges_written": 0, "reason": str(e)}
        finally:
            try:
                await self.workbook.close_session(
                    item_id, session_id, site_id=site_id, drive_id=drive_id
                )
            except Exception as e:  # best-effort; the writes already landed
                logger.warning("Failed to close session %s: %s", session_id, e)

    async def _apply(self, item_id, drive_id, site_id, session_id, plan) -> None:
        """Apply plan ops in a safe order within the probed session.

        Order: renames (so subsequent name-based addressing matches the proposed
        names) -> adds (+ their content) -> deletes -> reorders -> cell writes.
        """
        common = dict(
            item_id=item_id, drive_id=drive_id, site_id=site_id,
            session_id=session_id, auto_session=False, check_lock=False,
        )

        for op in plan.renames:
            await self.workbook.update_worksheet(sheet=op.old, name=op.new, **common)

        for op in plan.adds:
            await self.workbook.add_worksheet(name=op.name, **common)
            for w in op.writes:
                await self.workbook.update_range(
                    sheet=w.sheet, address=w.address, values=w.values, **common
                )

        for op in plan.deletes:
            await self.workbook.delete_worksheet(sheet=op.name, **common)

        for op in plan.reorders:
            await self.workbook.update_worksheet(
                sheet=op.name, position=op.position, **common
            )

        for w in plan.range_writes:
            await self.workbook.update_range(
                sheet=w.sheet, address=w.address, values=w.values, **common
            )
