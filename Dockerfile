# Image officielle Python slim (Debian Trixie)
FROM python:3.11-slim

# Métadonnées
LABEL org.opencontainers.image.title="Export du TCO" \
      org.opencontainers.image.description="Consolidation automatique des DPGF et remplissage du TCO" \
      org.opencontainers.image.version="2.2.0"

WORKDIR /app

# curl uniquement — requis pour le HEALTHCHECK
# (streamlit/pandas/openpyxl sont distribués sous forme de wheels, pas de compilation C)
RUN apt-get update \
    && apt-get install -y --no-install-recommends curl \
    && rm -rf /var/lib/apt/lists/*

# Dépendances Python — couche séparée pour profiter du cache Docker
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Code applicatif
COPY . .

EXPOSE 8501

# Utilisateur non-root pour réduire la surface d'attaque en production
RUN useradd -m -u 1000 tco && chown -R tco:tco /app
USER tco

HEALTHCHECK --interval=30s --timeout=10s --start-period=15s --retries=3 \
    CMD curl --fail --silent http://localhost:8501/_stcore/health || exit 1

ENTRYPOINT ["streamlit", "run", "app.py", \
            "--server.port=8501", \
            "--server.address=0.0.0.0", \
            "--server.headless=true"]
