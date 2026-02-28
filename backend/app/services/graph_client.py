import httpx
from datetime import datetime, timedelta
from typing import Optional, Any
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select
from msal import ConfidentialClientApplication

from app.config import get_settings
from app.models import Auth
from app.crypto import encrypt_token, decrypt_token
from app import audit

settings = get_settings()

# MSAL auto-adds these scopes, so we filter them from user-provided scopes
RESERVED_SCOPES = {"openid", "profile", "offline_access"}


def get_user_scopes() -> list[str]:
    return [s for s in settings.scopes if s.lower() not in RESERVED_SCOPES]


class GraphClient:
    def __init__(self, db: AsyncSession, auth: Auth):
        self.db = db
        self.auth = auth
        self.base_url = settings.graph_base_url
        self._client: Optional[httpx.AsyncClient] = None

    async def _get_access_token(self) -> str:
        # Check if token needs refresh (5-minute buffer)
        if datetime.utcnow() >= self.auth.expires_at - timedelta(minutes=5):
            await self._refresh_token()
        return decrypt_token(self.auth.access_token)

    async def _refresh_token(self) -> None:
        msal_app = ConfidentialClientApplication(
            settings.azure_client_id,
            authority=settings.authority,
            client_credential=settings.azure_client_secret,
        )

        refresh_token = decrypt_token(self.auth.refresh_token)
        result = msal_app.acquire_token_by_refresh_token(
            refresh_token,
            scopes=get_user_scopes(),
        )

        if "access_token" not in result:
            error_msg = result.get('error_description', 'Unknown error')
            audit.log_token_refresh(self.auth.email, success=False, error=error_msg)
            raise Exception(f"Token refresh failed: {error_msg}")

        self.auth.access_token = encrypt_token(result["access_token"])
        if "refresh_token" in result:
            self.auth.refresh_token = encrypt_token(result["refresh_token"])
        self.auth.expires_at = datetime.utcnow() + timedelta(seconds=result.get("expires_in", 3600))

        self.db.add(self.auth)
        await self.db.commit()

        audit.log_token_refresh(self.auth.email, success=True)

    async def _get_client(self) -> httpx.AsyncClient:
        if self._client is None:
            self._client = httpx.AsyncClient(timeout=30.0)
        return self._client

    async def _get_headers(self) -> dict:
        token = await self._get_access_token()
        return {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

    async def get(self, endpoint: str, params: Optional[dict] = None, extra_headers: Optional[dict] = None) -> dict:
        client = await self._get_client()
        headers = await self._get_headers()
        if extra_headers:
            headers.update(extra_headers)
        url = f"{self.base_url}{endpoint}"
        response = await client.get(url, headers=headers, params=params)
        response.raise_for_status()
        return response.json()

    async def post(self, endpoint: str, data: Optional[dict] = None) -> dict:
        client = await self._get_client()
        headers = await self._get_headers()
        url = f"{self.base_url}{endpoint}"
        response = await client.post(url, headers=headers, json=data)
        response.raise_for_status()
        if response.status_code == 204:
            return {}
        return response.json() if response.content else {}

    async def patch(self, endpoint: str, data: dict) -> dict:
        client = await self._get_client()
        headers = await self._get_headers()
        url = f"{self.base_url}{endpoint}"
        response = await client.patch(url, headers=headers, json=data)
        response.raise_for_status()
        if response.status_code == 204:
            return {}
        return response.json() if response.content else {}

    async def put(self, endpoint: str, content: bytes, content_type: str = "application/octet-stream") -> dict:
        client = await self._get_client()
        token = await self._get_access_token()
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": content_type,
        }
        url = f"{self.base_url}{endpoint}"
        response = await client.put(url, headers=headers, content=content)
        response.raise_for_status()
        return response.json() if response.content else {}

    async def delete(self, endpoint: str) -> None:
        client = await self._get_client()
        headers = await self._get_headers()
        url = f"{self.base_url}{endpoint}"
        response = await client.delete(url, headers=headers)
        response.raise_for_status()

    async def get_raw(self, endpoint: str) -> bytes:
        client = await self._get_client()
        headers = await self._get_headers()
        url = f"{self.base_url}{endpoint}"
        response = await client.get(url, headers=headers, follow_redirects=True)
        response.raise_for_status()
        return response.content

    async def close(self):
        if self._client:
            await self._client.aclose()
            self._client = None
