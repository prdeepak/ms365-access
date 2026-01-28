from fastapi import APIRouter, Depends, HTTPException, status
from fastapi.responses import RedirectResponse
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select, delete
from datetime import datetime, timedelta
from msal import ConfidentialClientApplication

from app.config import get_settings
from app.database import get_db
from app.models import Auth
from app.schemas import AuthStatus
from app.crypto import encrypt_token, decrypt_token
from app import audit

router = APIRouter(prefix="/auth", tags=["auth"])
settings = get_settings()

# MSAL auto-adds these scopes, so we filter them from user-provided scopes
RESERVED_SCOPES = {"openid", "profile", "offline_access"}


def get_user_scopes() -> list[str]:
    return [s for s in settings.scopes if s.lower() not in RESERVED_SCOPES]


def get_msal_app() -> ConfidentialClientApplication:
    return ConfidentialClientApplication(
        settings.azure_client_id,
        authority=settings.authority,
        client_credential=settings.azure_client_secret,
    )


@router.get("/login")
async def login():
    msal_app = get_msal_app()

    auth_url = msal_app.get_authorization_request_url(
        scopes=get_user_scopes(),
        redirect_uri=settings.azure_redirect_uri,
    )

    return RedirectResponse(url=auth_url)


@router.get("/callback")
async def callback(code: str, db: AsyncSession = Depends(get_db)):
    msal_app = get_msal_app()

    result = msal_app.acquire_token_by_authorization_code(
        code,
        scopes=get_user_scopes(),
        redirect_uri=settings.azure_redirect_uri,
    )

    if "error" in result:
        error_msg = result.get('error_description', result.get('error'))
        audit.log_login_attempt("unknown", success=False, error=error_msg)
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Authentication failed: {error_msg}",
        )

    if "access_token" not in result:
        audit.log_login_attempt("unknown", success=False, error="No access token received")
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail="No access token received",
        )

    # Get user info from id_token claims
    id_token_claims = result.get("id_token_claims", {})
    email = id_token_claims.get("preferred_username") or id_token_claims.get("email", "unknown@user.com")

    # Calculate expiration
    expires_in = result.get("expires_in", 3600)
    expires_at = datetime.utcnow() + timedelta(seconds=expires_in)

    # Clear existing auth records and insert new one
    await db.execute(delete(Auth))

    auth = Auth(
        email=email,
        access_token=encrypt_token(result["access_token"]),
        refresh_token=encrypt_token(result.get("refresh_token", "")),
        expires_at=expires_at,
    )
    db.add(auth)
    await db.commit()

    audit.log_login_attempt(email, success=True)
    return {"message": "Authentication successful", "email": email}


@router.get("/status", response_model=AuthStatus)
async def auth_status(db: AsyncSession = Depends(get_db)):
    result = await db.execute(select(Auth).limit(1))
    auth = result.scalar_one_or_none()

    if not auth:
        return AuthStatus(authenticated=False)

    return AuthStatus(
        authenticated=True,
        email=auth.email,
        expires_at=auth.expires_at,
    )


@router.post("/logout")
async def logout(db: AsyncSession = Depends(get_db)):
    # Get email before clearing for audit log
    result = await db.execute(select(Auth).limit(1))
    auth = result.scalar_one_or_none()
    email = auth.email if auth else "unknown"

    await db.execute(delete(Auth))
    await db.commit()

    audit.log_logout(email)
    return {"message": "Logged out successfully"}
