# MS365 Access - FastAPI Backend
# Pinned to a digest so `docker pull` is deterministic and tamper-evident (the
# floating python:3.11-slim tag can be silently repointed upstream). Re-pin when
# you intentionally upgrade Python or pull in base-image security patches:
#   docker buildx imagetools inspect python:3.11-slim --format '{{.Manifest.Digest}}'
FROM python:3.11-slim@sha256:f9fa7f851e38bfb19c9de3afbc4b86ae7176ea7aaf94535c31df5458d5849457

# Set working directory
WORKDIR /app

# Install system dependencies
RUN apt-get update && apt-get install -y \
    sqlite3 \
    && rm -rf /var/lib/apt/lists/*

# Set environment variables
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1

# Copy the hash-pinned lock first for layer caching
COPY backend/requirements.lock .

# Install Python dependencies, verifying every package against its pinned hash
RUN pip install --no-cache-dir --require-hashes -r requirements.lock

# Copy application code
COPY backend/ .

# Create data directory for SQLite
RUN mkdir -p /app/data

# Expose port
EXPOSE 8365

# Run the application
CMD ["uvicorn", "app.main:app", "--host", "0.0.0.0", "--port", "8365", "--reload"]
