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
