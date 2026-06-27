# CI smoke-test spec — ms365-access

**Goal:** add a no-credentials CI job + a minimal pytest smoke suite so a green check means
"the API app and the MCP server still import, the FastAPI route table builds, `/health`
responds, and dependencies resolve." Open a PR; do NOT merge.

**Authored by MARVIN (orchestrator). You are the local implementing agent.**

## Hard constraints
- Work on a **new branch** (`ci/add-pytest-smoke`). Never push `main`. Never merge.
- Do **not** read, print, move, or commit any `.env` or secret. (None exists in a clean
  clone — keep it that way.)
- Backend-only repo — **no frontend job**. Stop at `gh pr create`.

## Repo facts (verified — these are the gotchas that break a naive CI copy)
- `app/` package layout. **`backend/app/main.py` calls `get_settings()` at module import
  time** (line ~20). `app/config.py` `Settings` has **required, default-less** fields:
  `azure_client_id`, `azure_client_secret`, `azure_tenant_id`, `secret_key`. So
  `import app.main` **raises pydantic ValidationError without those env vars** — CI must
  inject dummy values or import fails.
- **Two lock/req pairs**: `requirements.lock`/`requirements.txt` (API) and
  `requirements-mcp.lock`/`requirements-mcp.txt` (MCP). CI should install the one(s) needed
  to import what the smoke covers (at minimum the API lock for `app.main`; add the MCP lock
  if you also smoke `mcp_server.py`).
- There is also a top-level `backend/mcp_server.py` (the MCP entrypoint).

## Deliverables
1. **`backend/tests/conftest.py`** — set the four dummy env vars **before any app import**
   (module top-level, not in a fixture, so collection-time imports see them):
   `AZURE_CLIENT_ID=ci-not-real`, `AZURE_CLIENT_SECRET=ci-not-real`,
   `AZURE_TENANT_ID=ci-not-real`, `SECRET_KEY=ci-test-secret-not-used-for-real-auth`.
   (ms365 already has 1 test — `test_workbook_diff.py`; integrate, don't clobber it.)
2. **`backend/tests/test_smoke.py`** — assert, with NO real credentials:
   - `import app.main` succeeds (proves settings + routers import under dummy env).
   - the FastAPI `app` builds its route table (assert `app.routes` is non-empty; `/health`
     is registered).
   - **`/health` via `fastapi.testclient.TestClient(app)` returns 200** — read the handler
     (`main.py` ~line 98) first; if it touches Graph/creds, mock that call or assert the real
     degraded behavior. Assert what the code actually does; don't force 200 if it can't.
   - (optional, if MCP lock installed) `import mcp_server` succeeds and tool schemas serialize.
3. **`.github/workflows/ci.yml`** — backend-only:
   - triggers: `push: branches:[main]` and `pull_request`.
   - `working-directory: backend`, `actions/setup-python@v5` py3.11, pip cache on the lock(s).
   - install: `pip install --require-hashes -r requirements.lock` (+ `-r requirements-mcp.lock`
     if smoking the MCP side).
   - run: `python -m pytest -q`.
   - **env: all four dummy vars above** (single `SECRET_KEY` is NOT enough for this repo).
4. Ensure `pytest` is available in CI (see slack-access spec §4 for the lock-regen vs
   `requirements-dev.txt` choice; same rule — let `uv.toml` govern seasoning, no CLI
   `--exclude-newer`). Note the choice in the PR.

## Acceptance before opening the PR
- In the clone: venv, install the lock(s) with `--require-hashes`, install pytest per (4),
  export the four dummy env vars, confirm `python -m pytest -q` **passes locally**. Do not
  open a non-draft PR if tests fail.

## PR
- Title: `ci: add no-creds pytest smoke + CI workflow`.
- Body: what was added, that tests pass locally with no credentials, how the dual-lock +
  import-time-settings gotchas were handled, the pytest-availability choice, and
  "for review — do not auto-merge."
