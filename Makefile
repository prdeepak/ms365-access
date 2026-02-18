# MS365 Access - Development Commands
# Ports: API=8365, MCP (openclaw)=8367

.PHONY: help up down build restart shell logs logs-mcp auth status check-docker gen-client check-client test

# Default target
help:
	@echo "MS365 Access - Available Commands"
	@echo ""
	@echo "Docker Commands:"
	@echo "  make up        - Start all containers (API + openclaw MCP server)"
	@echo "  make down      - Stop and remove all containers"
	@echo "  make build     - Build Docker images"
	@echo "  make restart   - Rebuild and restart all containers"
	@echo "  make logs      - Tail logs from all containers"
	@echo "  make logs-mcp  - Tail logs from openclaw MCP server only"
	@echo ""
	@echo "Shell Access:"
	@echo "  make shell     - Open shell in API container"
	@echo ""
	@echo "App Commands:"
	@echo "  make auth      - Open login page in browser"
	@echo "  make status    - Show app and auth status"
	@echo ""
	@echo "Testing:"
	@echo "  make test         - Run all checks (includes check-client)"
	@echo ""
	@echo "Code Generation:"
	@echo "  make gen-client   - Regenerate Python client + README from OpenAPI spec"
	@echo "  make check-client - Verify generated files are up-to-date (for CI)"
	@echo ""
	@echo "Ports: API=8365, MCP (openclaw, streamable-http)=8367"

# Check if Docker is running (supports OrbStack)
check-docker:
	@if ! docker info > /dev/null 2>&1; then \
		echo "Docker is not running. Starting OrbStack..."; \
		open -a OrbStack; \
		echo "Waiting for Docker to start..."; \
		while ! docker info > /dev/null 2>&1; do \
			sleep 1; \
		done; \
		echo "Docker is ready."; \
	fi

# Docker commands
up: check-docker
	docker compose up -d
	@echo ""
	@echo "Services started:"
	@echo "  API:      http://localhost:8365"
	@echo "  Docs:     http://localhost:8365/docs"
	@echo "  Auth:     http://localhost:8365/auth/login"
	@echo "  MCP:      http://localhost:8367/mcp  (openclaw tier, streamable-http)"

down:
	docker compose down

build: check-docker
	docker compose build

restart: down build up

logs:
	docker compose logs -f

logs-mcp:
	docker compose logs -f mcp

# Shell access
shell:
	docker compose exec api bash

# App commands
auth:
	open "http://localhost:8365/auth/login"

# Code generation
gen-client:
	@echo "Generating Python client and README API docs from OpenAPI spec..."
	python3 scripts/gen_client.py
	@echo ""

check-client:
	@python3 scripts/gen_client.py --check

test: check-client
	@echo "All checks passed."

status:
	@echo "Container status:"
	@docker compose ps 2>/dev/null || echo "Not running"
	@echo ""
	@echo "App status:"
	@curl -s http://localhost:8365/ 2>/dev/null | python3 -m json.tool || echo "API not responding"
	@echo "---"
	@curl -s http://localhost:8365/auth/status 2>/dev/null | python3 -m json.tool || echo "Auth not available"
	@echo "---"
	@echo "MCP server (openclaw):"
	@curl -s -o /dev/null -w "  http://localhost:8367  →  %{http_code}\n" http://localhost:8367/ 2>/dev/null || echo "  MCP not responding"
