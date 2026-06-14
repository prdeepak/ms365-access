"""No-credentials smoke tests.

A green run proves: the API app and its routers import under dummy settings,
the FastAPI route table builds, and `/health` responds 200 — without any real
Azure/Graph credentials. Dummy env vars are set in conftest.py at import time.
"""

from fastapi.testclient import TestClient


def test_import_app_main():
    """Importing app.main succeeds, proving settings + all routers import."""
    import app.main  # noqa: F401

    assert app.main.app is not None


def test_route_table_built():
    """The FastAPI app builds a non-empty route table including /health."""
    from app.main import app

    assert app.routes, "expected a non-empty route table"
    paths = {getattr(route, "path", None) for route in app.routes}
    assert "/health" in paths


def test_health_endpoint_returns_200():
    """`/health` returns 200. The handler just returns a static status dict
    (app/main.py ~line 98) — no Graph/credential access — so we assert the
    real behavior directly.

    The app installs TrustedHostMiddleware with allowed_hosts defaulting to
    ["localhost", "127.0.0.1"], so we point the client at http://localhost (a
    real, allowed Host header). TestClient's default "testserver" host is a test
    artifact that the middleware legitimately rejects with 400. TestClient is
    used without a context manager so the lifespan (init_db) does not run;
    /health has no such dependency."""
    from app.main import app

    client = TestClient(app, base_url="http://localhost")
    response = client.get("/health")

    assert response.status_code == 200
    assert response.json() == {"status": "healthy"}


def test_health_rejects_untrusted_host():
    """Sanity check on the security middleware: an untrusted Host header is
    rejected with 400, confirming the 200 above comes from a genuinely allowed
    host rather than the middleware being absent."""
    from app.main import app

    client = TestClient(app, base_url="http://evil.example.com")
    response = client.get("/health")

    assert response.status_code == 400
