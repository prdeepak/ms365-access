from fastapi import APIRouter, Depends, Query, UploadFile, File, Response
from typing import Optional

from app.dependencies import get_graph_client, get_current_auth, require_permission
from app.services.graph_client import GraphClient
from app.services.onedrive_service import OneDriveService
from app.models import Auth
from app.schemas import CreateFolderRequest, RenameItemRequest
from app import audit

router = APIRouter(prefix="/files", tags=["files"])


def get_onedrive_service(graph_client: GraphClient = Depends(get_graph_client)) -> OneDriveService:
    return OneDriveService(graph_client)


@router.get("/drives", dependencies=[Depends(require_permission("read:files"))])
async def list_drives(onedrive_service: OneDriveService = Depends(get_onedrive_service)):
    result = await onedrive_service.list_drives()
    return result.get("value", [])


@router.get("/drive/root", dependencies=[Depends(require_permission("read:files"))])
async def get_drive_root(
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    return await onedrive_service.get_drive_root(drive_id=drive_id)


@router.get("/items/{item_id}", dependencies=[Depends(require_permission("read:files"))])
async def get_item(
    item_id: str,
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    return await onedrive_service.get_item(item_id=item_id, drive_id=drive_id)


@router.get("/items/{item_id}/children", dependencies=[Depends(require_permission("read:files"))])
async def list_children(
    item_id: str,
    drive_id: Optional[str] = None,
    top: int = Query(100, ge=1, le=200),
    skip: int = Query(0, ge=0),
    order_by: str = "name",
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    result = await onedrive_service.list_children(
        item_id=item_id,
        drive_id=drive_id,
        top=top,
        skip=skip,
        order_by=order_by,
    )
    return {
        "items": result.get("value", []),
        "next_link": result.get("@odata.nextLink"),
    }


@router.get("/items/{item_id}/content", dependencies=[Depends(require_permission("read:files"))])
async def download_content(
    item_id: str,
    drive_id: Optional[str] = None,
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
    auth: Auth = Depends(get_current_auth),
):
    content = await onedrive_service.download_content(item_id=item_id, drive_id=drive_id)
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
    onedrive_service: OneDriveService = Depends(get_onedrive_service),
):
    result = await onedrive_service.search(query=q, drive_id=drive_id, top=top)
    return {
        "items": result.get("value", []),
        "next_link": result.get("@odata.nextLink"),
    }
