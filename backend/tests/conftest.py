"""Pytest configuration for the backend test suite.

`app/main.py` calls `get_settings()` at *module import time*, and
`app/config.py` `Settings` declares required, default-less fields
(`azure_client_id`, `azure_client_secret`, `azure_tenant_id`, `secret_key`).
Importing `app.main` therefore raises a pydantic ValidationError unless those
env vars exist. Set dummy values here at module top level (NOT inside a fixture)
so they are in place before pytest collection triggers any `import app.*`.

These are throwaway placeholders. They are NOT real credentials and are never
used for real authentication — they only satisfy settings validation so the app
and routers can be imported and the route table can be built in CI.
"""

import os

os.environ.setdefault("AZURE_CLIENT_ID", "ci-not-real")
os.environ.setdefault("AZURE_CLIENT_SECRET", "ci-not-real")
os.environ.setdefault("AZURE_TENANT_ID", "ci-not-real")
os.environ.setdefault("SECRET_KEY", "ci-test-secret-not-used-for-real-auth")
