"""
logger.py — Journalisation centralisée avec rotation de fichiers.

ARCH-4 : Remplace tous les print() par des appels logger structurés.
Le logger principal est accessible via get_logger().
"""

import logging
import logging.handlers
import os
import sys


def setup_logger(log_dir="logs", log_filename="tco_automator.log",
                 max_bytes=5 * 1024 * 1024, backup_count=3,
                 level="INFO"):
    """
    Configure le logger principal avec :
      - RotatingFileHandler : rotation à max_bytes, backup_count fichiers
      - StreamHandler       : sortie console (niveau WARNING et plus)

    Args:
        log_dir      : dossier où écrire le fichier log
        log_filename : nom du fichier log
        max_bytes    : taille max avant rotation
        backup_count : nombre de fichiers de backup à conserver
        level        : niveau de log (DEBUG/INFO/WARNING/ERROR)

    Returns:
        logging.Logger
    """
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, log_filename)

    logger = logging.getLogger("tco_automator")
    if logger.handlers:
        # déjà initialisé (rechargements Streamlit)
        return logger

    logger.setLevel(getattr(logging, level.upper(), logging.INFO))

    fmt = logging.Formatter(
        "%(asctime)s | %(levelname)-8s | %(name)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )

    # --- Fichier avec rotation ---
    try:
        fh = logging.handlers.RotatingFileHandler(
            log_path,
            maxBytes=max_bytes,
            backupCount=backup_count,
            encoding="utf-8",
        )
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except OSError as e:
        # Ne pas bloquer l'app si le log ne peut pas être créé
        print(f"[WARN] Impossible d'ouvrir {log_path}: {e}", file=sys.stderr)

    # --- Console (WARNING+) ---
    ch = logging.StreamHandler(sys.stderr)
    ch.setLevel(logging.WARNING)
    ch.setFormatter(fmt)
    logger.addHandler(ch)

    return logger


def get_logger(name=None):
    """
    Retourne un logger nommé, enfant du logger principal.

    Usage:
        from logger import get_logger
        log = get_logger(__name__)
        log.info("Fichier parsé")
        log.warning("Code non trouvé : 01.1A")
        log.error("Erreur lecture Excel", exc_info=True)
    """
    parent = logging.getLogger("tco_automator")
    if not parent.handlers:
        # auto-setup si pas encore initialisé
        setup_logger()
    return logging.getLogger(f"tco_automator.{name}") if name else parent
