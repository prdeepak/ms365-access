import hashlib
import json
import logging
from datetime import datetime
from typing import Callable

from fastapi import Depends, HTTPException, status
from fastapi.security import HTTPAuthorizationCredentials, HTTPBearer
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_db
from app.models import ApiKey, Auth
from app.services.graph_client import GraphClient

logger = logging.getLogger(__name__)

# All valid permissions in the system
VALID_PERMISSIONS = {
    "read:mail",
    "read:calendar",
    "read:contacts",
    "read:files",
    "write:draft",
    "write:mail",
    "write:calendar",
    "write:contacts",
    "write:files",
    "admin",
}

# Named permission tiers — use when creating API keys with POST /api-keys {"tier": "..."}
# Each tier is a minimal, expandable permission set. "admin" grants all permissions implicitly.
TIER_PERMISSIONS: dict[str, list[str]] = {
    "admin": ["admin"],
    "openclaw": [
        "read:mail",
        "write:draft",      # create/edit drafts; NOT write:mail (no send/reply/move/delete)
        "read:calendar",
        "write:calendar",
        "read:contacts",
        "write:contacts",
    ],
}

bearer_scheme = HTTPBearer()


async def verify_api_key(
    credentials: HTTPAuthorizationCredentials = Depends(bearer_scheme),
    db: AsyncSession = Depends(get_db),
) -> ApiKey:
    """Extract Bearer token, SHA256-hash it, look up in DB. Returns the ApiKey row."""
    raw_key = credentials.credentials
    key_hash = hashlib.sha256(raw_key.encode()).hexdigest()

    result = await db.execute(
        select(ApiKey).where(ApiKey.key_hash == key_hash, ApiKey.is_active == True)
    )
    api_key = result.scalar_one_or_none()

    if not api_key:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Invalid or inactive API key.",
            headers={"WWW-Authenticate": "Bearer"},
        )

    # Update last_used_at (fire-and-forget, non-blocking)
    api_key.last_used_at = datetime.utcnow()
    db.add(api_key)
    await db.commit()

    return api_key


def require_permission(permission: str) -> Callable:
    """Return a FastAPI dependency that enforces a specific permission.

    Usage in a route:
        @router.get("/mail/messages", dependencies=[Depends(require_permission("read:mail"))])

    The 'admin' permission implicitly grants access to everything.
    """
    if permission not in VALID_PERMISSIONS:
        raise ValueError(f"Unknown permission: {permission}")

    async def _check(api_key: ApiKey = Depends(verify_api_key)) -> ApiKey:
        perms = json.loads(api_key.permissions) if isinstance(api_key.permissions, str) else api_key.permissions
        if "admin" in perms or permission in perms:
            return api_key
        raise HTTPException(
            status_code=status.HTTP_403_FORBIDDEN,
            detail=f"API key '{api_key.name}' lacks required permission: {permission}",
        )

    return _check


# ---------- existing helpers (unchanged semantics) ----------


async def get_current_auth(db: AsyncSession = Depends(get_db)) -> Auth:
    result = await db.execute(select(Auth).limit(1))
    auth = result.scalar_one_or_none()

    if not auth:
        raise HTTPException(
            status_code=status.HTTP_401_UNAUTHORIZED,
            detail="Not authenticated. Please visit /auth/login to authenticate.",
        )

    return auth


async def get_graph_client(
    db: AsyncSession = Depends(get_db),
    auth: Auth = Depends(get_current_auth),
) -> GraphClient:
    return GraphClient(db, auth)
