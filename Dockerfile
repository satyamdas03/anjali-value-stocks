# syntax=docker/dockerfile:1

# AnjaliValueStocks — Automated Stock Screener
# Base image: Python 3.11 slim
FROM python:3.11-slim

# Install minimal system dependencies for lxml, numpy, and common wheels
RUN apt-get update && apt-get install -y --no-install-recommends \
    gcc \
    libxml2 \
    libxslt1.1 \
    && rm -rf /var/lib/apt/lists/*

# Set working directory inside the container
WORKDIR /app

# Install Python dependencies first (better Docker layer caching)
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy all project files into the image
COPY . .

# Optional environment variables (documented; override at runtime if needed)
# TZ                — timezone for scheduler logs and APScheduler (default: UTC)
# PYTHONUNBUFFERED  — ensures Python stdout/stderr is unbuffered in containers
# OUTPUT_DIR        — if set, scripts may redirect outputs here (not yet implemented in legacy scripts)
ENV PYTHONUNBUFFERED=1
ENV PYTHONDONTWRITEBYTECODE=1

# Expose health check port
EXPOSE 8080

# Health check for Railway
HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 \
    CMD python -c "import urllib.request; urllib.request.urlopen('http://localhost:8080/health')" || exit 1

# Run the scheduler as the main long-running process
CMD ["python", "scheduler.py"]
