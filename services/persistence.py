"""
persistence.py — Sauvegarde et chargement des projets (format JSON sécurisé).

Remplace pickle par JSON + gzip pour éliminer le risque d'exécution de
code arbitraire. Les DataFrames sont sérialisés via .to_dict(orient="records").
"""

import gzip
import json
import os

import pandas as pd

from config import PROJECTS_DIR, TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)

# Extension des fichiers projet (nouvelle version)
PROJECT_EXT = ".tco.json.gz"
# Ancienne extension (pickle) pour migration
LEGACY_EXT = ".tco"


def _project_path(name: str) -> str:
    """Retourne le chemin complet du fichier projet."""
    return os.path.join(PROJECTS_DIR, f"{name}{PROJECT_EXT}")


def _legacy_path(name: str) -> str:
    """Retourne le chemin de l'ancien format pickle."""
    return os.path.join(PROJECTS_DIR, f"{name}{LEGACY_EXT}")


def save_project(name: str, session_state) -> tuple[bool, str]:
    """
    Sauvegarde l'état actuel dans un fichier JSON compressé.

    Args:
        name: nom du projet
        session_state: st.session_state contenant les données

    Returns:
        (success, message)
    """
    if not name:
        return False, "Le nom du projet est vide."

    os.makedirs(PROJECTS_DIR, exist_ok=True)
    path = _project_path(name)

    # Sérialisation des DataFrames en dictionnaires
    tco_df = session_state.get("tco_df")
    merged_df = session_state.get("merged_df")

    company_data_serialized = {}
    for comp_name, comp_info in session_state.get("company_data", {}).items():
        company_data_serialized[comp_name] = {
            "dpgf_df": comp_info["dpgf_df"].to_dict(orient="records"),
            "parse_alerts": comp_info["parse_alerts"],
            "filename": comp_info["filename"],
        }

    data = {
        "version": 2,  # Version du format de sauvegarde
        "tco_df": tco_df.to_dict(orient="records") if tco_df is not None else None,
        "tco_meta": session_state.get("tco_meta"),
        "company_data": company_data_serialized,
        "step": session_state.get("step", 1),
        "all_alerts": session_state.get("all_alerts", []),
        "merged_df": merged_df.to_dict(orient="records") if merged_df is not None else None,
        "project_name": name,
        "tva_rate": session_state.get("tva_rate", TVA_DEFAULT),
    }

    try:
        with gzip.open(path, "wt", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, default=str)
        log.info("Projet sauvegardé (JSON) : %s", name)

        # Supprimer l'ancien fichier pickle s'il existe
        legacy = _legacy_path(name)
        if os.path.exists(legacy):
            try:
                os.remove(legacy)
                log.info("Ancien fichier pickle supprimé : %s", legacy)
            except OSError:
                pass

        return True, f"Projet '{name}' sauvegardé avec succès."
    except Exception as e:
        log.error("Erreur sauvegarde projet %s : %s", name, e)
        return False, f"Erreur technique : {e}"


def load_project(name: str, session_state) -> tuple[bool, str]:
    """
    Charge un projet depuis un fichier JSON compressé.

    Args:
        name: nom du projet
        session_state: st.session_state à restaurer

    Returns:
        (success, message)
    """
    path = _project_path(name)

    if not os.path.exists(path):
        # Tentative de migration depuis l'ancien format pickle
        legacy = _legacy_path(name)
        if os.path.exists(legacy):
            return _migrate_legacy_project(name, legacy, session_state)
        return False, "Le fichier de projet n'existe plus."

    try:
        with gzip.open(path, "rt", encoding="utf-8") as f:
            data = json.load(f)

        # Restauration des DataFrames
        tco_records = data.get("tco_df")
        session_state.tco_df = pd.DataFrame(tco_records) if tco_records else None

        merged_records = data.get("merged_df")
        session_state.merged_df = pd.DataFrame(merged_records) if merged_records else None

        session_state.tco_meta = data.get("tco_meta", {})
        session_state.step = data.get("step", 1)
        session_state.all_alerts = data.get("all_alerts", [])
        session_state.current_project = name
        session_state.tva_rate = data.get("tva_rate", TVA_DEFAULT)

        # Restauration des données entreprises
        company_data = {}
        for comp_name, comp_info in data.get("company_data", {}).items():
            company_data[comp_name] = {
                "dpgf_df": pd.DataFrame(comp_info["dpgf_df"]),
                "parse_alerts": comp_info["parse_alerts"],
                "filename": comp_info["filename"],
            }
        session_state.company_data = company_data

        log.info("Projet chargé (JSON) : %s", name)
        return True, f"Projet '{name}' chargé."
    except Exception as e:
        log.error("Erreur chargement projet %s : %s", name, e)
        return False, f"Erreur de lecture : {e}"


def _migrate_legacy_project(name: str, legacy_path: str, session_state) -> tuple[bool, str]:
    """
    Tente de migrer un projet depuis l'ancien format pickle vers JSON.

    ⚠️ pickle.load est utilisé UNE DERNIÈRE FOIS pour la migration.
    Après migration, le fichier pickle est supprimé.
    """
    import pickle

    log.warning("Migration projet legacy (pickle) : %s", name)
    try:
        with open(legacy_path, "rb") as f:
            data = pickle.load(f)  # noqa: S301 — migration uniquement

        # Restaurer dans le session_state
        session_state.tco_df = data.get("tco_df")
        session_state.company_data = data.get("company_data", {})
        session_state.tco_meta = data.get("tco_meta", {})
        session_state.step = data.get("step", 1)
        session_state.all_alerts = data.get("all_alerts", [])
        session_state.merged_df = data.get("merged_df")
        session_state.current_project = name

        # Re-sauvegarder immédiatement en JSON
        ok, msg = save_project(name, session_state)
        if ok:
            log.info("Migration réussie pour %s — pickle supprimé", name)
            return True, f"Projet '{name}' migré et chargé."
        return True, f"Projet '{name}' chargé (migration partielle)."
    except Exception as e:
        log.error("Erreur migration projet %s : %s", name, e)
        return False, f"Erreur de lecture de l'ancien format : {e}"


def list_projects() -> list[str]:
    """Liste les noms de projets disponibles (JSON et legacy)."""
    if not os.path.exists(PROJECTS_DIR):
        return []

    names: set[str] = set()

    for f in os.listdir(PROJECTS_DIR):
        if f.endswith(PROJECT_EXT):
            # Retirer l'extension composée .tco.json.gz
            name = f.removesuffix(PROJECT_EXT)
            names.add(name)
        elif f.endswith(LEGACY_EXT):
            # Ancien format pickle
            names.add(os.path.splitext(f)[0])

    return sorted(names)


def delete_project(name: str) -> bool:
    """Supprime un fichier projet (JSON et/ou legacy)."""
    deleted = False

    for ext in (PROJECT_EXT, LEGACY_EXT):
        path = os.path.join(PROJECTS_DIR, f"{name}{ext}")
        if os.path.exists(path):
            try:
                os.remove(path)
                log.info("Projet supprimé : %s (%s)", name, ext)
                deleted = True
            except OSError as e:
                log.warning("Impossible de supprimer %s : %s", path, e)

    return deleted
