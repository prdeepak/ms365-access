"""Excel Workbook API router.

Cell/range/table-level writes to a live Excel file, co-authoring-safe.
See app.services.workbook_service for the why/how.
"""

from fastapi import APIRouter, Depends, HTTPException, Query
from typing import Optional

from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.workbook_service import WorkbookService, WorkbookLockedError
from app.models import Auth
from app import audit

router = APIRouter(prefix="/workbook", tags=["workbook"])


def get_workbook_service(
    graph_client: GraphClient = Depends(get_graph_client),
) -> WorkbookService:
    return WorkbookService(graph_client)


@router.get("/items/{item_id}/lock-state", dependencies=[Depends(require_permission("read:files"))])
async def lock_state(
    item_id: str,
    site_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Report whether the file is checked out (co-authoring is not a checkout)."""
    return await workbook_service.get_lock_state(item_id=item_id, site_id=site_id)


@router.post("/items/{item_id}/session", dependencies=[Depends(require_permission("write:files"))])
async def create_session(
    item_id: str,
    site_id: Optional[str] = None,
    persist: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Create a workbook session. Use the returned id for subsequent calls."""
    try:
        return await workbook_service.create_session(
            item_id=item_id, site_id=site_id, persist=persist,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))


@router.delete("/items/{item_id}/session", dependencies=[Depends(require_permission("write:files"))])
async def close_session(
    item_id: str,
    session_id: str,
    site_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Close a workbook session."""
    return await workbook_service.close_session(
        item_id=item_id, session_id=session_id, site_id=site_id,
    )


@router.get("/items/{item_id}/worksheets", dependencies=[Depends(require_permission("read:files"))])
async def list_worksheets(
    item_id: str,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """List worksheets in the workbook."""
    result = await workbook_service.list_worksheets(
        item_id=item_id, site_id=site_id, session_id=session_id,
    )
    return result.get("value", [])


@router.get("/items/{item_id}/tables", dependencies=[Depends(require_permission("read:files"))])
async def list_tables(
    item_id: str,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """List tables in the workbook."""
    result = await workbook_service.list_tables(
        item_id=item_id, site_id=site_id, session_id=session_id,
    )
    return result.get("value", [])


@router.get("/items/{item_id}/range", dependencies=[Depends(require_permission("read:files"))])
async def get_range(
    item_id: str,
    sheet: str,
    address: str,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Read a cell range, e.g. sheet='Sheet1' address='A1:C10'."""
    return await workbook_service.get_range(
        item_id=item_id, sheet=sheet, address=address,
        site_id=site_id, session_id=session_id,
    )


@router.patch("/items/{item_id}/range", dependencies=[Depends(require_permission("write:files"))])
async def update_range(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Write a 2D array of values into a range.

    Body: {"sheet": "Sheet1", "address": "A1:B2", "values": [["x", 1], ["y", 2]]}

    With no session_id and auto_session=True (default), the server opens and
    closes a persistent session around the write for a clean one-call edit.
    """
    sheet = body.get("sheet")
    address = body.get("address")
    values = body.get("values")
    if not sheet or not address or values is None:
        raise HTTPException(
            status_code=400,
            detail="'sheet', 'address', and 'values' are required",
        )
    try:
        result = await workbook_service.update_range(
            item_id=item_id, sheet=sheet, address=address, values=values,
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_range:{item_id}", f"{sheet}!{address}")
    return result


@router.post("/items/{item_id}/tables/{table}/rows", dependencies=[Depends(require_permission("write:files"))])
async def add_table_row(
    item_id: str,
    table: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Append rows to a table. Body: {"values": [["east", "pear", 4]]}.

    Auto-sessions around the write when no session_id is given (see update_range).
    """
    values = body.get("values")
    if values is None:
        raise HTTPException(status_code=400, detail="'values' is required")
    try:
        result = await workbook_service.add_table_row(
            item_id=item_id, table=table, values=values,
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_table_row:{item_id}", table)
    return result


# -- Worksheet management ------------------------------------------------


@router.get("/items/{item_id}/worksheet", dependencies=[Depends(require_permission("read:files"))])
async def get_worksheet(
    item_id: str,
    sheet: str,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Read a single worksheet's metadata (id, name, position, visibility)."""
    return await workbook_service.get_worksheet(
        item_id=item_id, sheet=sheet, site_id=site_id, session_id=session_id,
    )


@router.get("/items/{item_id}/used-range", dependencies=[Depends(require_permission("read:files"))])
async def get_used_range(
    item_id: str,
    sheet: str,
    values_only: bool = False,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    workbook_service: WorkbookService = Depends(get_workbook_service),
):
    """Read the worksheet's used range. values_only excludes formatting-only cells."""
    return await workbook_service.get_used_range(
        item_id=item_id, sheet=sheet, values_only=values_only,
        site_id=site_id, session_id=session_id,
    )


@router.post("/items/{item_id}/worksheets", dependencies=[Depends(require_permission("write:files"))])
async def add_worksheet(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Add a worksheet. Body: {"name": "Sheet2"} (name optional)."""
    try:
        result = await workbook_service.add_worksheet(
            item_id=item_id, name=body.get("name"),
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_add_sheet:{item_id}", body.get("name") or "")
    return result


@router.delete("/items/{item_id}/worksheet", dependencies=[Depends(require_permission("write:files"))])
async def delete_worksheet(
    item_id: str,
    sheet: str,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Delete a worksheet by name or id (sheet given as a query param)."""
    try:
        result = await workbook_service.delete_worksheet(
            item_id=item_id, sheet=sheet,
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_delete(auth.email, f"workbook_sheet:{item_id}:{sheet}")
    return result


@router.patch("/items/{item_id}/worksheet", dependencies=[Depends(require_permission("write:files"))])
async def update_worksheet(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Rename / reorder / show-hide a worksheet.

    Body: {"sheet": "Sheet1", "name": "New", "position": 2, "visibility": "Hidden"}.
    `sheet` is required; at least one of name/position/visibility must be set.
    """
    sheet = body.get("sheet")
    name = body.get("name")
    position = body.get("position")
    visibility = body.get("visibility")
    if not sheet:
        raise HTTPException(status_code=400, detail="'sheet' is required")
    if name is None and position is None and visibility is None:
        raise HTTPException(
            status_code=400,
            detail="one of 'name', 'position', or 'visibility' is required",
        )
    try:
        result = await workbook_service.update_worksheet(
            item_id=item_id, sheet=sheet, name=name, position=position,
            visibility=visibility, site_id=site_id, session_id=session_id,
            auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_update_sheet:{item_id}", sheet)
    return result


@router.post("/items/{item_id}/worksheet/copy", dependencies=[Depends(require_permission("write:files"))])
async def copy_worksheet(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Copy a worksheet.

    Body: {"sheet": "Sheet1", "name": "Copy", "position_type": "After",
    "relative_to": "Sheet1"}. `sheet` is required.
    """
    sheet = body.get("sheet")
    if not sheet:
        raise HTTPException(status_code=400, detail="'sheet' is required")
    try:
        result = await workbook_service.copy_worksheet(
            item_id=item_id, sheet=sheet, name=body.get("name"),
            position_type=body.get("position_type"), relative_to=body.get("relative_to"),
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_copy_sheet:{item_id}", sheet)
    return result


@router.post("/items/{item_id}/worksheet/protect", dependencies=[Depends(require_permission("write:files"))])
async def protect_worksheet(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Protect a worksheet. Body: {"sheet": "Sheet1", "options": {...}} (options optional)."""
    sheet = body.get("sheet")
    if not sheet:
        raise HTTPException(status_code=400, detail="'sheet' is required")
    try:
        result = await workbook_service.protect_worksheet(
            item_id=item_id, sheet=sheet, options=body.get("options"),
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_protect_sheet:{item_id}", sheet)
    return result


@router.post("/items/{item_id}/worksheet/unprotect", dependencies=[Depends(require_permission("write:files"))])
async def unprotect_worksheet(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Remove protection from a worksheet. Body: {"sheet": "Sheet1"}."""
    sheet = body.get("sheet")
    if not sheet:
        raise HTTPException(status_code=400, detail="'sheet' is required")
    try:
        result = await workbook_service.unprotect_worksheet(
            item_id=item_id, sheet=sheet,
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_unprotect_sheet:{item_id}", sheet)
    return result


@router.post("/items/{item_id}/range/clear", dependencies=[Depends(require_permission("write:files"))])
async def clear_range(
    item_id: str,
    body: dict,
    site_id: Optional[str] = None,
    session_id: Optional[str] = None,
    auto_session: bool = True,
    workbook_service: WorkbookService = Depends(get_workbook_service),
    auth: Auth = Depends(get_current_auth),
):
    """Clear a cell range.

    Body: {"sheet": "Sheet1", "address": "A1:B2", "apply_to": "All"}.
    `apply_to` is 'All' | 'Formats' | 'Contents' (default 'All').
    """
    sheet = body.get("sheet")
    address = body.get("address")
    if not sheet or not address:
        raise HTTPException(status_code=400, detail="'sheet' and 'address' are required")
    try:
        result = await workbook_service.clear_range(
            item_id=item_id, sheet=sheet, address=address,
            apply_to=body.get("apply_to", "All"),
            site_id=site_id, session_id=session_id, auto_session=auto_session,
        )
    except WorkbookLockedError as e:
        raise HTTPException(status_code=409, detail=str(e))
    audit.log_file_upload(auth.email, f"workbook_clear_range:{item_id}", f"{sheet}!{address}")
    return result
