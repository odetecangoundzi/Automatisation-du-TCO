"""
config.py — Configuration centralisée du TCO Automator.

ARCH-3 : Toutes les constantes métier et techniques sont ici.
Modifier ce fichier pour adapter l'application à votre contexte.
Les valeurs peuvent être surchargées via un fichier .env (voir .env.example).
# Trigger CI: test-run-001
"""

import os

from dotenv import load_dotenv

load_dotenv()  # Charge les variables depuis .env si le fichier existe

# ---------------------------------------------------------------------------
# Répertoires
# ---------------------------------------------------------------------------

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
OUTPUT_DIR = os.path.join(BASE_DIR, "outputs")
LOG_DIR    = os.path.join(BASE_DIR, "logs")
PROJECTS_DIR = os.path.join(BASE_DIR, "projects")

# ---------------------------------------------------------------------------
# Fichiers
# ---------------------------------------------------------------------------

ALLOWED_EXTENSIONS = {".xlsx", ".xlsm", ".xls", ".xlsb"}
MAX_FILE_SIZE_MB   = 20          # Taille max par fichier uploadé

# ---------------------------------------------------------------------------
# Entreprises
# ---------------------------------------------------------------------------

MAX_COMPANIES = 100               # Nombre max d'entreprises simultanées
COMPANY_NAME_MAX_LEN = 60        # Longueur max du nom d'entreprise

# ---------------------------------------------------------------------------
# Métier — TVA
# ---------------------------------------------------------------------------

TVA_OPTIONS = {
    "5,5 %": 0.055,   # Travaux de rénovation résidentielle
    "10 %":  0.10,    # Travaux de rénovation non HLM
    "20 %":  0.20,    # Taux normal
}
TVA_DEFAULT = 0.20

# ---------------------------------------------------------------------------
# Métier — Alertes
# ---------------------------------------------------------------------------

# Tolérance pour l'incohérence de total Qu × PU vs Px_Tot_HT
TOTAL_TOLERANCE_ABS = 0.10    # max 10 centimes d'écart absolu
TOTAL_TOLERANCE_REL = 0.001   # OU 0.1 % d'écart relatif

# ---------------------------------------------------------------------------
# Logging
# ---------------------------------------------------------------------------

LOG_FILENAME       = "tco_automator.log"
LOG_MAX_BYTES      = 5 * 1024 * 1024   # 5 MB par fichier
LOG_BACKUP_COUNT   = 3                  # Garder 3 fichiers de rotation
LOG_LEVEL          = "INFO"             # DEBUG | INFO | WARNING | ERROR

# ---------------------------------------------------------------------------
# Application
# ---------------------------------------------------------------------------

APP_TITLE   = "TCO Automator"
APP_VERSION = "2.1.0"
APP_ICON    = "📊"
