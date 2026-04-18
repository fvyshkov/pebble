FROM python:3.12-slim

# Install Node.js (for frontend build) and curl
RUN apt-get update && apt-get install -y --no-install-recommends \
      curl ca-certificates gnupg \
    && curl -fsSL https://deb.nodesource.com/setup_20.x | bash - \
    && apt-get install -y --no-install-recommends nodejs \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Python deps (cached separately from source)
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt uvicorn

# Frontend deps (cached separately from source)
COPY frontend/package.json frontend/package-lock.json* ./frontend/
RUN cd frontend && (npm ci || npm install)

# Copy the rest and build frontend
COPY . .
RUN cd frontend && npm run build

# Pebble backend serves the built frontend from /frontend/dist
# Data (pebble.db) lives in /data so it can be mounted to a persistent disk.
ENV PEBBLE_DB=/data/pebble.db
RUN mkdir -p /data

EXPOSE 8000
CMD ["sh", "-c", "uvicorn backend.main:app --host 0.0.0.0 --port ${PORT:-8000}"]
