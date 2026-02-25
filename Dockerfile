# Utiliser une image Python officielle légère
FROM python:3.11-slim

# Définir le répertoire de travail
WORKDIR /app

# Installer les dépendances système nécessaires
RUN apt-get update && apt-get install -y \
    build-essential \
    curl \
    software-properties-common \
    git \
    && rm -rf /var/lib/apt/lists/*

# Copier le fichier des dépendances
COPY requirements.txt .

# Installer les dépendances Python
RUN pip3 install --no-cache-dir -r requirements.txt

# Copier tout le code du projet
COPY . .

# Exposer le port par défaut de Streamlit (8501)
EXPOSE 8501

# Santé de l'application
HEALTHCHECK CMD curl --fail http://localhost:8501/_stcore/health

# Lancer l'application
ENTRYPOINT ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
