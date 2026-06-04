import json

from fastapi import APIRouter, Depends, Form, HTTPException, Query, UploadFile, File, Response
from typing import Optional

from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.onedrive_service import OneDriveService
from app.services.workbook_service import WorkbookService
from app.services.smart_update_service import SmartUpdateService
from app.models import Auth
from app.schemas import CreateFolderRequest, RenameItemRequest
from app import audit

router = APIRouter(prefix="/files", tags=["files"])


def get_onedrive_service(graph_client: GraphClient = Depends(get_graph_client)) -> OneDriveService:
    return OneDriveService(graph_client)


def get_smart_update_service(
    graph_client: GraphClient = Depends(get_graph_client),
) -> SmartUpdateService:
    return SmartUpdateService(
        OneDriveService(graph_client), WorkbookService(graph_client)
    )


@router.get("/drives", dependencies=[Depends(require_permission("read:files"))])
async def list_drives(
    user: Optional[str] = Query(None, description="UPN of another user (e.g. caroline@revivalgourmet.com) whose drives you have access to."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    result = await onedrive_service.list_drives(user=user)
    return result.get("value", [])


@router.get("/drive/root", dependencies=[Depends(require_permission("read:files"))])
async def get_drive_root(
    drive_id: Optional[str] = None,
    user: Optional[str] = Query(None, description="UPN of another user whose default OneDrive root you have access to."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    return await onedrive_service.get_drive_root(drive_id=drive_id, user=user)


@router.get("/items/{item_id}", dependencies=[Depends(require_permission("read:files"))])
async def get_item(
    item_id: str,
    drive_id: Optional[str] = None,
    user: Optional[str] = Query(None, description="UPN of another user whose OneDrive item you have access to."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    return await onedrive_service.get_item(item_id=item_id, drive_id=drive_id, user=user)


@router.get("/items/{item_id}/children", dependencies=[Depends(require_permission("read:files"))])
async def list_children(
    item_id: str,
    drive_id: Optional[str] = None,
    top: int = Query(100, ge=1, le=200),
    skip: int = Query(0, ge=0),
    order_by: str = "name",
    user: Optional[str] = Query(None, description="UPN of another user whose OneDrive folder you have access to."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    result = await onedrive_service.list_children(
        item_id=item_id,
        drive_id=drive_id,
        top=top,
        skip=skip,
        order_by=order_by,
        user=user,
    )
    return result.get("value", [])


@router.get("/items/{item_id}/content", dependencies=[Depends(require_permission("read:files"))])
async def download_content(
    item_id: str,
    drive_id: Optional[str] = None,
    user: Optional[str] = Query(None, description="UPN of another user whose OneDrive file you have access to."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
    auth: Auth = Depends(get_current_auth),
):
    content = await onedrive_service.download_content(item_id=item_id, drive_id=drive_id, user=user)
    audit.log_file_download(auth.email, item_id)
    return Response(content=content, media_type="application/octet-stream")


@router.put("/items/{parent_id}:/{filename}:/content", dependencies=[Depends(require_permission("write:files"))])
async def upload_content(
    parent_id: str,
    filename: str,
    file: UploadFile = File(...),
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
    auth: Auth = Depends(get_current_auth),
):
    content = await file.read()
    content_type = file.content_type or "application/octet-stream"

    result = await onedrive_service.upload_content(
        parent_id=parent_id,
        filename=filename,
        content=content,
        content_type=content_type,
        drive_id=drive_id,
    )
    audit.log_file_upload(auth.email, filename, parent_id)
    return result


@router.put("/items/{item_id}/content", dependencies=[Depends(require_permission("write:files"))])
async def replace_content(
    item_id: str,
    file: UploadFile = File(...),
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
    auth: Auth = Depends(get_current_auth),
):
    """Replace the content of an existing file (keeps the same item ID).

    OneDrive/SharePoint preserves version history, so the previous content
    is still accessible via the file's version history.
    """
    content = await file.read()
    content_type = file.content_type or "application/octet-stream"

    result = await onedrive_service.replace_content(
        item_id=item_id,
        content=content,
        content_type=content_type,
        drive_id=drive_id,
    )
    audit.log_file_upload(auth.email, f"replace:{item_id}", item_id)
    return result


@router.post("/items/{item_id}/smart-update", dependencies=[Depends(require_permission("write:files"))])
async def smart_update(
    item_id: str,
    file: UploadFile = File(...),
    drive_id: Optional[str] = None,
    site_id: Optional[str] = None,
    region_map: Optional[str] = Form(None),
    smart_update_service: SmartUpdateService = Depends(get_smart_update_service),
    auth: Auth = Depends(get_current_auth),
):
    """Replace a SharePoint/OneDrive .xlsx in place, with a live-edit fallback.

    Tries a whole-file replace first. If the file is open in Excel (423 Locked),
    it diffs the proposed workbook against the live one and — when every change
    is a value/formula/structure edit it can reproduce via a Graph session
    (values/formulas inside declared `region_map` regions; worksheet
    add/delete/rename/reorder) — applies them surgically, leaving formatting
    untouched. Otherwise it returns `deferred` so the caller can ask the user to
    close the file and retry a clean replace.

    `region_map` is an optional JSON object, e.g.
    `{"AP payments": {"data": "A2:U200"}}`. Without it, most diffs defer (safe
    default). Returns `{"mode", "ranges_written", "reason"}` where mode is one of
    `replaced` | `live-edited` | `deferred`.
    """
    content = await file.read()
    try:
        rmap = json.loads(region_map) if region_map else None
    except json.JSONDecodeError as e:
        raise HTTPException(status_code=400, detail=f"region_map is not valid JSON: {e}")

    result = await smart_update_service.smart_update(
        item_id=item_id,
        new_bytes=content,
        drive_id=drive_id,
        site_id=site_id,
        region_map=rmap,
    )
    audit.log_file_upload(auth.email, f"smart_update:{item_id}", result["mode"])
    return result


@router.delete("/items/{item_id}", dependencies=[Depends(require_permission("write:files"))])
async def delete_item(
    item_id: str,
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
    auth: Auth = Depends(get_current_auth),
):
    await onedrive_service.delete_item(item_id=item_id, drive_id=drive_id)
    audit.log_file_delete(auth.email, item_id)
    return {"message": "Item deleted successfully"}


@router.post("/items/{parent_id}/folder", dependencies=[Depends(require_permission("write:files"))])
async def create_folder(
    parent_id: str,
    request: CreateFolderRequest,
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    return await onedrive_service.create_folder(
        parent_id=parent_id,
        folder_name=request.name,
        drive_id=drive_id,
    )


@router.patch("/items/{item_id}", dependencies=[Depends(require_permission("write:files"))])
async def update_item(
    item_id: str,
    request: RenameItemRequest,
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    if request.parent_id:
        return await onedrive_service.move_item(
            item_id=item_id,
            new_parent_id=request.parent_id,
            new_name=request.name,
            drive_id=drive_id,
        )
    elif request.name:
        return await onedrive_service.rename_item(
            item_id=item_id,
            new_name=request.name,
            drive_id=drive_id,
        )
    return {"message": "No changes requested"}


@router.get("/search", dependencies=[Depends(require_permission("read:files"))])
async def search_files(
    q: str,
    drive_id: Optional[str] = None,
    top: int = Query(25, ge=1, le=100),
    user: Optional[str] = Query(None, description="UPN of another user whose OneDrive you have access to. Default: your own OneDrive."),
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    result = await onedrive_service.search(query=q, drive_id=drive_id, top=top, user=user)
    return result.get("value", [])
