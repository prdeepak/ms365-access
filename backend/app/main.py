import logging
from contextlib import asynccontextmanager
from fastapi import FastAPI, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
from starlette.middleware.trustedhost import TrustedHostMiddleware
import httpx

from app.config import get_settings
from app.database import init_db
from app.routers import auth, mail, calendar, files, sharepoint

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s",
)
logger = logging.getLogger(__name__)

settings = get_settings()


@asynccontextmanager
async def lifespan(app: FastAPI):
    # Startup
    logger.info("Starting MS365 Access API")
    await init_db()
    yield
    # Shutdown
    logger.info("Shutting down MS365 Access API")


app = FastAPI(
    title="MS365 Access API",
    description="API for accessing Microsoft 365 services (Email, Calendar, OneDrive)",
    version="1.0.0",
    lifespan=lifespan,
)

# Security: Validate Host header to prevent host header injection attacks
app.add_middleware(
    TrustedHostMiddleware,
    allowed_hosts=settings.allowed_hosts,
)

# Security: Explicit CORS configuration instead of wildcards
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.cors_origins,
    allow_credentials=True,
    allow_methods=["GET", "POST", "PATCH", "PUT", "DELETE", "OPTIONS"],
    allow_headers=["Content-Type", "Authorization", "Accept"],
)


# Exception handler for HTTP errors from Graph API
@app.exception_handler(httpx.HTTPStatusError)
async def http_exception_handler(request: Request, exc: httpx.HTTPStatusError):
    try:
        error_detail = exc.response.json()
    except Exception:
        error_detail = {"message": str(exc)}

    return JSONResponse(
        status_code=exc.response.status_code,
        content={
            "error": "Graph API Error",
            "detail": error_detail,
        },
    )


# Include routers
app.include_router(auth.router)
app.include_router(mail.router)
app.include_router(calendar.router)
app.include_router(files.router)
app.include_router(sharepoint.router)


@app.get("/")
async def root():
    return {
        "name": "MS365 Access API",
        "version": "1.0.0",
        "docs": "/docs",
        "status": "running",
    }


@app.get("/health")
async def health():
    return {"status": "healthy"}


if __name__ == "__main__":
    import uvicorn

    # Default binds to localhost for security; use APP_HOST=0.0.0.0 for Docker
    uvicorn.run(app, host=settings.app_host, port=settings.app_port)
