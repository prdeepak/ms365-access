from fastapi import APIRouter, Depends, HTTPException, Query, Response, UploadFile, File
from typing import Optional

from app.dependencies import get_graph_client, get_current_auth
from app.services.graph_client import GraphClient
from app.services.sharepoint_service import SharePointService
from app.models import Auth
from app import audit

router = APIRouter(prefix="/sharepoint", tags=["sharepoint"])


def get_sharepoint_service(
    graph_client: GraphClient = Depends(get_graph_client),
) -> SharePointService:
    return SharePointService(graph_client)


@router.get("/sites/{host_path:path}")
async def resolve_site(
    host_path: str,
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """Resolve a SharePoint site by hostname/path.

    Example: /sharepoint/sites/contoso.sharepoint.com/sites/Finance
    """
    return await sharepoint_service.resolve_site(host_path)


@router.get("/drives")
async def list_drives(
    site_id: str,
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """List document libraries (drives) for a SharePoint site."""
    result = await sharepoint_service.list_drives(site_id)
    return result.get("value", [])


@router.get("/items/{item_id}/children")
async def list_children(
    item_id: str,
    site_id: str,
    top: int = Query(100, ge=1, le=200),
    order_by: str = "name",
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """List children of a folder in a SharePoint drive."""
    result = await sharepoint_service.list_children(
        site_id=site_id,
        item_id=item_id,
        top=top,
        order_by=order_by,
    )
    return result.get("value", [])


@router.get("/items/{item_id}/content")
async def download_content(
    item_id: str,
    site_id: str,
    format: Optional[str] = None,
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
    auth: Auth = Depends(get_current_auth),
):
    """Download file content from SharePoint. Optionally convert format (e.g. format=pdf)."""
    content = await sharepoint_service.download_content(
        site_id=site_id, item_id=item_id, format=format
    )
    audit.log_file_download(auth.email, item_id)

    media_type = "application/octet-stream"
    if format == "pdf":
        media_type = "application/pdf"

    return Response(content=content, media_type=media_type)


@router.put("/items/{parent_id}:/{filename}:/content")
async def upload_content(
    parent_id: str,
    filename: str,
    site_id: str,
    file: UploadFile = File(...),
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
    auth: Auth = Depends(get_current_auth),
):
    """Upload a file to a SharePoint site's default drive.

    Uses /sites/{site_id}/drive path to avoid b!-prefixed drive IDs.
    """
    content = await file.read()
    content_type = file.content_type or "application/octet-stream"
    result = await sharepoint_service.upload_content(
        site_id=site_id,
        parent_id=parent_id,
        filename=filename,
        content=content,
        content_type=content_type,
    )
    audit.log_file_upload(auth.email, filename, parent_id)
    return result


@router.put("/items/{item_id}/content")
async def replace_content(
    item_id: str,
    site_id: str,
    file: UploadFile = File(...),
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
    auth: Auth = Depends(get_current_auth),
):
    """Replace the content of an existing file in SharePoint (keeps the same item ID).

    SharePoint preserves version history, so the previous content is still
    accessible via the file's version history.
    """
    content = await file.read()
    content_type = file.content_type or "application/octet-stream"
    result = await sharepoint_service.replace_content(
        site_id=site_id,
        item_id=item_id,
        content=content,
        content_type=content_type,
    )
    audit.log_file_upload(auth.email, f"replace:{item_id}", item_id)
    return result


@router.get("/items/{item_id}")
async def get_item(
    item_id: str,
    site_id: str,
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """Get item metadata from a SharePoint drive."""
    return await sharepoint_service.get_item(site_id=site_id, item_id=item_id)


@router.get("/search")
async def search(
    q: str,
    site_id: str,
    top: int = Query(25, ge=1, le=100),
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """Search within a SharePoint site's default drive."""
    result = await sharepoint_service.search(site_id=site_id, query=q, top=top)
    return result.get("value", [])


@router.get("/resolve")
async def resolve_url(
    url: str,
    sharepoint_service: SharePointService = Depends(get_sharepoint_service),
):
    """Resolve a SharePoint sharing URL to item metadata + site_id + item_id."""
    try:
        return await sharepoint_service.resolve_sharepoint_url(url)
    except ValueError as e:
        raise HTTPException(status_code=404, detail=str(e))
