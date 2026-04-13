"""Power BI Service — read-only access to Power BI REST API.

Uses a separate token audience (analysis.windows.net/powerbi/api) from
MS Graph. Acquires Power BI tokens via the existing MSAL refresh token.
"""

import logging
from datetime import datetime, timedelta
from typing import Optional

import httpx
from msal import ConfidentialClientApplication
from sqlalchemy.ext.asyncio import AsyncSession

from app.config import get_settings
from app.models import Auth
from app.crypto import decrypt_token
from app import audit

logger = logging.getLogger(__name__)
settings = get_settings()

POWERBI_BASE_URL = "https://api.powerbi.com/v1.0/myorg"
POWERBI_SCOPES = [
    "https://analysis.windows.net/powerbi/api/Dataset.ReadWrite.All",
    "https://analysis.windows.net/powerbi/api/Workspace.Read.All",
    "https://analysis.windows.net/powerbi/api/Report.ReadWrite.All",
]


class PowerBIService:
    """Power BI REST API client.

    Acquires access tokens for the Power BI resource using the stored
    MSAL refresh token, separate from the MS Graph token.
    """

    def __init__(self, db: AsyncSession, auth: Auth):
        self.db = db
        self.auth = auth
        self._pbi_token: Optional[str] = None
        self._pbi_token_expires: Optional[datetime] = None
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_pbi_token(self) -> str:
        """Acquire (or return cached) Power BI access token."""
        now = datetime.utcnow()
        if self._pbi_token and self._pbi_token_expires and now < self._pbi_token_expires:
            return self._pbi_token

        msal_app = ConfidentialClientApplication(
            settings.azure_client_id,
            authority=settings.authority,
            client_credential=settings.azure_client_secret,
        )

        refresh_token = decrypt_token(self.auth.refresh_token)
        result = msal_app.acquire_token_by_refresh_token(
            refresh_token,
            scopes=POWERBI_SCOPES,
        )

        if "access_token" not in result:
            error_msg = result.get("error_description", result.get("error", "Unknown error"))
            logger.error("Power BI token acquisition failed: %s", error_msg)
            raise Exception(f"Power BI token acquisition failed: {error_msg}")

        self._pbi_token = result["access_token"]
        self._pbi_token_expires = now + timedelta(
            seconds=result.get("expires_in", 3600) - 300  # 5-min buffer
        )

        logger.info("Acquired Power BI token (expires in %ss)", result.get("expires_in", "?"))
        return self._pbi_token

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None:
            self._client = httpx.AsyncClient(timeout=30.0)
        return self._client

    async def _get(self, endpoint: str, params: Optional[dict] = None) -> dict:
        token = await self._get_pbi_token()
        client = await self._get_client()
        url = f"{POWERBI_BASE_URL}{endpoint}"
        response = await client.get(
            url,
            headers={"Authorization": f"Bearer {token}"},
            params=params,
        )
        response.raise_for_status()
        return response.json()

    async def _post(self, endpoint: str, data: Optional[dict] = None) -> dict:
        token = await self._get_pbi_token()
        client = await self._get_client()
        url = f"{POWERBI_BASE_URL}{endpoint}"
        response = await client.post(
            url,
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            json=data,
            timeout=60.0,  # DAX queries can be slow
        )
        if response.status_code >= 400:
            # Try to extract a descriptive Power BI error message
            try:
                err_body = response.json()
                err_detail = err_body.get("error", {})
                err_code = err_detail.get("code", "")
                err_msg = err_detail.get("message", "")
                # Some errors nest detail further
                pbi_detail = err_detail.get("pbi.error", {}).get("details", [])
                detail_msgs = [d.get("detail", {}).get("value", "") for d in pbi_detail if d.get("detail")]
                full_msg = err_msg or err_code
                if detail_msgs:
                    full_msg = f"{full_msg} — {'; '.join(detail_msgs)}"
                if err_code == "DatasetExecuteQueriesError" and not err_msg:
                    full_msg = (
                        "DAX executeQueries failed. This usually means the workspace "
                        "does not have Power BI Premium or Premium Per User capacity."
                    )
                if full_msg:
                    logger.error("Power BI API %s %s: %s", response.status_code, endpoint, full_msg)
                    raise httpx.HTTPStatusError(
                        f"Power BI error ({response.status_code}): {full_msg}",
                        request=response.request,
                        response=response,
                    )
            except httpx.HTTPStatusError:
                raise
            except (ValueError, KeyError):
                pass
            response.raise_for_status()
        # 202 Accepted (e.g. refresh trigger) returns empty body
        if response.status_code == 202 or not response.content:
            return {}
        return response.json()

    # ------------------------------------------------------------------
    # Public API
    # ------------------------------------------------------------------

    async def list_workspaces(self) -> list[dict]:
        """List all Power BI workspaces (groups) the user has access to."""
        result = await self._get("/groups")
        return result.get("value", [])

    async def list_datasets(self, workspace_id: str) -> list[dict]:
        """List datasets in a workspace."""
        result = await self._get(f"/groups/{workspace_id}/datasets")
        return result.get("value", [])

    async def list_tables(self, workspace_id: str, dataset_id: str) -> list[dict]:
        """List tables in a dataset (push datasets only — standard datasets
        use DAX INFO functions instead)."""
        try:
            result = await self._get(
                f"/groups/{workspace_id}/datasets/{dataset_id}/tables"
            )
            return result.get("value", [])
        except httpx.HTTPStatusError as e:
            if e.response.status_code in (400, 404):
                # Standard (non-push) datasets don't support /tables — use DAX instead
                logger.info("Dataset %s is not a push dataset, falling back to DAX INFO.TABLES()", dataset_id)
                return await self._list_tables_via_dax(workspace_id, dataset_id)
            raise

    async def _list_tables_via_dax(
        self, workspace_id: str, dataset_id: str
    ) -> list[dict]:
        """Discover tables via DAX INFO.TABLES() for standard datasets.

        Falls back to returning a helpful error if DAX queries aren't
        available (requires Premium or PPU capacity).
        """
        try:
            dax = "EVALUATE INFO.TABLES()"
            result = await self.execute_query(workspace_id, dataset_id, dax)
            rows = result.get("rows", [])
            return [{"name": r.get("[Name]", r.get("Name", "?")), "source": "dax"} for r in rows]
        except Exception as e:
            logger.warning("DAX INFO.TABLES() failed for dataset %s: %s", dataset_id, e)
            return [{
                "error": "Cannot list tables for this dataset",
                "detail": (
                    "This is a standard (non-push) dataset. The REST /tables endpoint "
                    "only works for push datasets, and DAX queries require Power BI "
                    "Premium or Premium Per User capacity. To enable table discovery, "
                    "either: (1) assign the workspace to Premium/PPU capacity, or "
                    "(2) use the Power BI web UI to inspect the dataset schema."
                ),
            }]

    async def execute_query(
        self, workspace_id: str, dataset_id: str, dax_query: str
    ) -> dict:
        """Execute a DAX query against a dataset.

        Returns {"columns": [...], "rows": [...]} from the first result table.
        Requires Power BI Premium or Premium Per User capacity.
        """
        payload = {
            "queries": [{"query": dax_query}],
            "serializerSettings": {"includeNulls": True},
        }
        result = await self._post(
            f"/groups/{workspace_id}/datasets/{dataset_id}/executeQueries",
            data=payload,
        )

        # Parse the response — results.tables[0].rows
        tables = result.get("results", [{}])[0].get("tables", [])
        if not tables:
            return {"columns": [], "rows": []}

        table = tables[0]
        return {
            "columns": table.get("columns", []),
            "rows": table.get("rows", []),
        }

    async def list_reports(self, workspace_id: str) -> list[dict]:
        """List reports in a workspace."""
        result = await self._get(f"/groups/{workspace_id}/reports")
        return result.get("value", [])

    async def trigger_refresh(self, workspace_id: str, dataset_id: str) -> dict:
        """Trigger a dataset refresh.

        Returns immediately with {"status": "triggered"} on success (202).
        Power BI Pro: 8 refreshes/day max. Premium: 48/day.
        """
        await self._post(
            f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes",
            data={},
        )
        return {"status": "triggered"}

    async def list_refreshes(
        self, workspace_id: str, dataset_id: str, top: int = 10
    ) -> list[dict]:
        """List recent dataset refresh history.

        Returns refresh entries with: requestId, id, refreshType,
        startTime, endTime, status, serviceExceptionJson.
        """
        result = await self._get(
            f"/groups/{workspace_id}/datasets/{dataset_id}/refreshes",
            params={"$top": top},
        )
        return result.get("value", [])

    async def close(self):
        if self._client:
            await self._client.aclose()
            self._client = None
