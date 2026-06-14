#!/usr/bin/env bash
# Rebuild .venv-api from the hash-pinned lock after the fastapi/starlette upgrade.
# Safe to re-run. Does NOT touch .venv-mcp (unaffected by this upgrade).
set -euo pipefail

cd "$(dirname "$0")"   # backend/

echo "==> Pulling latest main..."
git pull --ff-only

echo "==> Rebuilding .venv-api from requirements.lock (hash-pinned)..."
rm -rf .venv-api
uv venv .venv-api --python 3.11
uv pip install --python .venv-api/bin/python --require-hashes -r requirements.lock

echo "==> Verifying starlette version in the new venv..."
.venv-api/bin/python -c "import starlette, fastapi; print(f'  fastapi {fastapi.__version__}, starlette {starlette.__version__}')"

echo
echo "Done. Now restart and verify:"
echo "  cd ~/marvin && make ms365-restart"
echo "  curl -s -o /dev/null -w '%{http_code}\\n' http://localhost:8365/health   # expect 200"
echo "  curl -s -o /dev/null -w '%{http_code}\\n' http://localhost:8367/          # expect 401"
