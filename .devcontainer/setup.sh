#!/bin/bash
set -e

# Backend setup
cd /workspace/backend
python -m venv .venv
.venv/bin/pip install --upgrade pip
.venv/bin/pip install -r requirements.txt

echo ""
echo "Setup complete!"
echo ""
echo "To start the server: cd backend && .venv/bin/uvicorn app.main:app --reload --port 8365"
echo ""
if [ ! -f /workspace/.env-mount/.env ]; then
  echo "WARNING: No .env file found. Create ~/.config/ms365-access/.env with your credentials."
  echo "         See .env.example for required variables."
fi
