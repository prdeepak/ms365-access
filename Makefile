# MS365 Access - Development Commands
# Port: API=8365

.PHONY: help up down build restart shell logs auth status check-docker

# Default target
help:
	@echo "MS365 Access - Available Commands"
	@echo ""
	@echo "Docker Commands:"
	@echo "  make up        - Start container (checks Docker first)"
	@echo "  make down      - Stop and remove container"
	@echo "  make build     - Build Docker image"
	@echo "  make restart   - Rebuild and restart container"
	@echo "  make logs      - Tail logs from container"
	@echo ""
	@echo "Shell Access:"
	@echo "  make shell     - Open shell in API container"
	@echo ""
	@echo "App Commands:"
	@echo "  make auth      - Open login page in browser"
	@echo "  make status    - Show app and auth status"
	@echo ""
	@echo "Port: API=8365"

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
	@echo "Service started:"
	@echo "  API: http://localhost:8365"
	@echo "  Docs: http://localhost:8365/docs"
	@echo "  Auth: http://localhost:8365/auth/login"

down:
	docker compose down

build: check-docker
	docker compose build

restart: down build up

logs:
	docker compose logs -f

# Shell access
shell:
	docker compose exec api bash

# App commands
auth:
	open "http://localhost:8365/auth/login"

status:
	@echo "Container status:"
	@docker compose ps 2>/dev/null || echo "Not running"
	@echo ""
	@echo "App status:"
	@curl -s http://localhost:8365/ 2>/dev/null | python3 -m json.tool || echo "API not responding"
	@echo "---"
	@curl -s http://localhost:8365/auth/status 2>/dev/null | python3 -m json.tool || echo "Auth not available"
