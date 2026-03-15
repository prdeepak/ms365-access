"""Power BI router — read-only endpoints for workspaces, datasets, and DAX queries."""

import logging

import httpx
from fastapi import APIRouter, Depends, Query, Body
from fastapi.responses import JSONResponse
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_db
from app.dependencies import get_current_auth, require_permission
from app.models import Auth
from app.services.powerbi_service import PowerBIService

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/powerbi", tags=["powerbi"])


def get_powerbi_service(
    db: AsyncSession = Depends(get_db),
    auth: Auth = Depends(get_current_auth),
) -> PowerBIService:
    return PowerBIService(db, auth)


# ------------------------------------------------------------------
# Workspaces
# ------------------------------------------------------------------


@router.get(
    "/workspaces",
    dependencies=[Depends(require_permission("read:powerbi"))],
)
async def list_workspaces(
    service: PowerBIService = Depends(get_powerbi_service),
):
    """List all Power BI workspaces the user has access to."""
    return await service.list_workspaces()


# ------------------------------------------------------------------
# Datasets
# ------------------------------------------------------------------


@router.get(
    "/workspaces/{workspace_id}/datasets",
    dependencies=[Depends(require_permission("read:powerbi"))],
)
async def list_datasets(
    workspace_id: str,
    service: PowerBIService = Depends(get_powerbi_service),
):
    """List datasets in a workspace."""
    return await service.list_datasets(workspace_id)


@router.get(
    "/workspaces/{workspace_id}/datasets/{dataset_id}/tables",
    dependencies=[Depends(require_permission("read:powerbi"))],
)
async def list_tables(
    workspace_id: str,
    dataset_id: str,
    service: PowerBIService = Depends(get_powerbi_service),
):
    """List tables in a dataset. Falls back to DAX INFO.TABLES() for standard datasets."""
    try:
        return await service.list_tables(workspace_id, dataset_id)
    except httpx.HTTPStatusError as e:
        return JSONResponse(
            {"error": "Power BI API Error", "detail": str(e)},
            status_code=e.response.status_code,
        )


# ------------------------------------------------------------------
# DAX Query
# ------------------------------------------------------------------


@router.post(
    "/workspaces/{workspace_id}/datasets/{dataset_id}/query",
    dependencies=[Depends(require_permission("read:powerbi"))],
)
async def execute_query(
    workspace_id: str,
    dataset_id: str,
    dax_query: str = Body(..., embed=True, description="DAX query string"),
    service: PowerBIService = Depends(get_powerbi_service),
):
    """Execute a DAX query against a dataset.

    Requires Power BI Premium or Premium Per User capacity.
    Returns columns and rows from the first result table.
    """
    try:
        return await service.execute_query(workspace_id, dataset_id, dax_query)
    except httpx.HTTPStatusError as e:
        return JSONResponse(
            {"error": "DAX query failed", "detail": str(e)},
            status_code=e.response.status_code,
        )


# ------------------------------------------------------------------
# Reports
# ------------------------------------------------------------------


@router.get(
    "/workspaces/{workspace_id}/reports",
    dependencies=[Depends(require_permission("read:powerbi"))],
)
async def list_reports(
    workspace_id: str,
    service: PowerBIService = Depends(get_powerbi_service),
):
    """List reports in a workspace."""
    return await service.list_reports(workspace_id)
