"""No-credentials smoke tests for the MCP server.

A green run proves `mcp_server.py` still imports, the FastMCP instance still
accepts the transport-security settings it is constructed with, the
streamable-http ASGI app still builds, and every `@mcp.tool()` still registers —
all against the exact dependency set `Dockerfile.mcp` installs.

Why this exists: the `backend-smoke` job installs `requirements.lock` +
`requirements-dev.lock`, and neither contains `mcp`. Before this file, an `mcp`
version bump could not break any CI job, because nothing in CI imported it. The
dedicated `mcp-smoke` job runs these tests against `requirements-mcp.lock`.

These tests are skipped (not failed) when `mcp` is absent, so the `backend-smoke`
job — which deliberately does not install it — stays green.
"""

import asyncio

import pytest

pytest.importorskip("mcp", reason="mcp is only installed in the mcp-smoke job")


def test_import_mcp_server():
    """Importing mcp_server succeeds, proving the FastMCP construction at module
    import time (including TransportSecuritySettings) is still valid API."""
    import mcp_server

    assert mcp_server.mcp is not None


def test_streamable_http_app_builds():
    """`streamable_http_app()` returns a mounted ASGI app.

    This is the entrypoint `main()` hands to uvicorn for the remote transport,
    so a signature or return-type change upstream would break the deployed
    server at startup rather than at import.
    """
    from mcp_server import mcp

    app = mcp.streamable_http_app()

    assert app is not None
    assert callable(app), "expected an ASGI-callable application"


def test_tools_register():
    """Every tool registers and exposes a name + input schema.

    An empty or partial tool list means a decorator or schema-generation change
    silently dropped tools — the server would still boot, but with no usable
    surface, so assert the list is non-empty and well-formed.
    """
    from mcp_server import mcp

    tools = asyncio.run(mcp.list_tools())

    assert tools, "expected a non-empty tool list"
    for tool in tools:
        assert tool.name, "every tool must have a name"
        assert tool.inputSchema is not None, f"{tool.name} is missing an input schema"

    names = {tool.name for tool in tools}
    # Spot-check one tool per service area so a regression that drops a whole
    # router's worth of tools fails loudly instead of shrinking the count.
    for expected in ("mail_send", "calendar_list_events", "files_search"):
        assert expected in names, f"expected tool {expected!r} to be registered"


def test_dns_rebinding_protection_configured():
    """The transport-security settings survive the FastMCP round-trip.

    This is the Host/Origin allowlist that backs the remote transport. Asserting
    on the live instance (not a freshly constructed one) catches an upstream
    refactor that renames or stops retaining the field — which would leave the
    server booting with the protection silently off.
    """
    from mcp_server import mcp

    security = mcp.settings.transport_security

    assert security is not None, "transport_security was not retained"
    assert security.enable_dns_rebinding_protection is True
    assert "127.0.0.1:*" in security.allowed_hosts
