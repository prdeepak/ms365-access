from fastapi import Depends, HTTPException, status
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.database import get_db
from app.models import Auth
from app.services.graph_client import GraphClient


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
