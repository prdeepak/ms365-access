import hashlib
import json
import logging
import secrets

from fastapi import APIRouter, Depends, HTTPException, status
from sqlalchemy import select
from sqlalchemy.ext.asyncio import AsyncSession

from app.database import get_db
from app.dependencies import require_permission, VALID_PERMISSIONS, TIER_PERMISSIONS
from app.models import ApiKey
from app.schemas import ApiKeyCreate, ApiKeyCreated, ApiKeyResponse, ApiKeyUpdate
from app import audit

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/api-keys", tags=["api-keys"])


def _parse_permissions(api_key: ApiKey) -> list[str]:
    return json.loads(api_key.permissions) if isinstance(api_key.permissions, str) else api_key.permissions


def _to_response(api_key: ApiKey) -> dict:
    return {
        "id": api_key.id,
        "name": api_key.name,
        "tier": api_key.tier,
        "permissions": _parse_permissions(api_key),
        "created_at": api_key.created_at,
        "last_used_at": api_key.last_used_at,
        "is_active": api_key.is_active,
    }


@router.get("", response_model=list[ApiKeyResponse])
async def list_api_keys(
    db: AsyncSession = Depends(get_db),
    _: ApiKey = Depends(require_permission("admin")),
):
    """List all API keys (admin only). Raw key values are never returned."""
    result = await db.execute(select(ApiKey).order_by(ApiKey.created_at.desc()))
    keys = result.scalars().all()
    return [_to_response(k) for k in keys]


@router.post("", response_model=ApiKeyCreated, status_code=status.HTTP_201_CREATED)
async def create_api_key(
    request: ApiKeyCreate,
    db: AsyncSession = Depends(get_db),
    _: ApiKey = Depends(require_permission("admin")),
):
    """Create a new API key. The raw key is returned only once.

    Supply either `tier` (e.g. "openclaw") to use a predefined permission set,
    or supply an explicit `permissions` list. Providing both is an error.
    """
    # Resolve permissions from tier or explicit list
    if request.tier and request.permissions:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="Provide either 'tier' or 'permissions', not both.",
        )
    if request.tier:
        if request.tier not in TIER_PERMISSIONS:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Unknown tier '{request.tier}'. Valid tiers: {', '.join(sorted(TIER_PERMISSIONS))}",
            )
        resolved_permissions = TIER_PERMISSIONS[request.tier]
    else:
        if not request.permissions:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail="Provide either 'tier' or a non-empty 'permissions' list.",
            )
        invalid = set(request.permissions) - VALID_PERMISSIONS
        if invalid:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid permissions: {', '.join(sorted(invalid))}",
            )
        resolved_permissions = request.permissions

    # Check for duplicate name
    existing = await db.execute(select(ApiKey).where(ApiKey.name == request.name))
    if existing.scalar_one_or_none():
        raise HTTPException(
            status_code=status.HTTP_409_CONFLICT,
            detail=f"An API key with name '{request.name}' already exists.",
        )

    raw_key = secrets.token_urlsafe(48)
    key_hash = hashlib.sha256(raw_key.encode()).hexdigest()

    api_key = ApiKey(
        key_hash=key_hash,
        name=request.name,
        tier=request.tier,
        permissions=json.dumps(sorted(resolved_permissions)),
    )
    db.add(api_key)
    await db.commit()
    await db.refresh(api_key)

    audit.log_event("api_keys", "create", details={"name": request.name})
    logger.info(f"API key created: {request.name}")

    resp = _to_response(api_key)
    resp["raw_key"] = raw_key
    return resp


@router.patch("/{key_id}", response_model=ApiKeyResponse)
async def update_api_key(
    key_id: int,
    request: ApiKeyUpdate,
    db: AsyncSession = Depends(get_db),
    _: ApiKey = Depends(require_permission("admin")),
):
    """Update an API key's name, permissions, or active status."""
    result = await db.execute(select(ApiKey).where(ApiKey.id == key_id))
    api_key = result.scalar_one_or_none()
    if not api_key:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="API key not found.")

    if request.name is not None:
        api_key.name = request.name
    if request.permissions is not None:
        invalid = set(request.permissions) - VALID_PERMISSIONS
        if invalid:
            raise HTTPException(
                status_code=status.HTTP_400_BAD_REQUEST,
                detail=f"Invalid permissions: {', '.join(sorted(invalid))}",
            )
        api_key.permissions = json.dumps(sorted(request.permissions))
    if request.is_active is not None:
        api_key.is_active = request.is_active

    db.add(api_key)
    await db.commit()
    await db.refresh(api_key)

    audit.log_event("api_keys", "update", details={"key_id": key_id})
    return _to_response(api_key)


@router.delete("/{key_id}")
async def revoke_api_key(
    key_id: int,
    db: AsyncSession = Depends(get_db),
    _: ApiKey = Depends(require_permission("admin")),
):
    """Revoke (deactivate) an API key. Does not delete it from the DB."""
    result = await db.execute(select(ApiKey).where(ApiKey.id == key_id))
    api_key = result.scalar_one_or_none()
    if not api_key:
        raise HTTPException(status_code=status.HTTP_404_NOT_FOUND, detail="API key not found.")

    api_key.is_active = False
    db.add(api_key)
    await db.commit()

    audit.log_event("api_keys", "revoke", details={"key_id": key_id, "name": api_key.name})
    return {"message": f"API key '{api_key.name}' has been revoked."}
