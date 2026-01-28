import asyncio
import json
import uuid
from datetime import datetime
from typing import Callable, Any
from sqlalchemy.ext.asyncio import AsyncSession
from sqlalchemy import select

from app.models import BackgroundJob
from app.database import async_session_maker


async def create_job(db: AsyncSession, job_type: str, total: int = 0) -> BackgroundJob:
    job = BackgroundJob(
        id=str(uuid.uuid4()),
        job_type=job_type,
        status="pending",
        progress=0,
        total=total,
    )
    db.add(job)
    await db.commit()
    await db.refresh(job)
    return job


async def update_job_progress(
    db: AsyncSession,
    job_id: str,
    progress: int,
    total: int,
    status: str = "running",
) -> None:
    """Update job progress using job_id to re-fetch the job in the current session."""
    result = await db.execute(select(BackgroundJob).where(BackgroundJob.id == job_id))
    job = result.scalar_one()
    job.progress = progress
    job.total = total
    job.status = status
    job.updated_at = datetime.utcnow()
    await db.commit()


async def complete_job(
    db: AsyncSession,
    job_id: str,
    total: int,
    result: Any = None,
    error: str = None,
) -> None:
    """Complete job using job_id to re-fetch the job in the current session."""
    db_result = await db.execute(select(BackgroundJob).where(BackgroundJob.id == job_id))
    job = db_result.scalar_one()
    job.status = "complete" if not error else "error"
    job.progress = total
    job.total = total
    job.result = json.dumps(result) if result else None
    job.error = error
    job.updated_at = datetime.utcnow()
    await db.commit()


async def run_batch_operation(
    job_id: str,
    items: list,
    operation: Callable,
) -> dict:
    """
    Run a batch operation with its own database session.

    This function creates a new database session that is independent of the
    request lifecycle, allowing it to run in a background task after the
    HTTP response has been sent.
    """
    async with async_session_maker() as db:
        total = len(items)

        # Update job to running status
        await update_job_progress(db, job_id, 0, total, "running")

        results = {"success": [], "failed": []}

        for i, item in enumerate(items):
            try:
                await operation(item)
                results["success"].append(item)
            except Exception as e:
                results["failed"].append({"item": item, "error": str(e)})

            await update_job_progress(db, job_id, i + 1, total)

        await complete_job(db, job_id, total, result=results)
        return results
