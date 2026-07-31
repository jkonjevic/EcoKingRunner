# EcoKing web UI — one image for Render, Koyeb, or plain Docker.
FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PYTHONDONTWRITEBYTECODE=1 \
    PIP_NO_CACHE_DIR=1 \
    PLAYWRIGHT_BROWSERS_PATH=/ms-playwright \
    ECOKING_MODE=cloud \
    DATA_DIR=/home/user/data \
    HOST=0.0.0.0 \
    PORT=8765

WORKDIR /app

# Chromium and its system libraries are the slow layer, so they are installed
# before the application code and cached across deploys.
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt \
    && playwright install --with-deps chromium \
    && chmod -R a+rX /ms-playwright \
    && rm -rf /var/lib/apt/lists/*

# Run as non-root; harmless on Render/Koyeb and required by some hosts.
RUN useradd --create-home --uid 1000 user \
    && mkdir -p /home/user/data \
    && chown -R user:user /home/user

COPY --chown=user:user . /app
USER user

# The platform overrides PORT at runtime (Render, Koyeb) — the app already
# reads it from the environment, so this EXPOSE is just documentation.
EXPOSE 8765
CMD ["python", "-m", "ecoking.webapp"]
