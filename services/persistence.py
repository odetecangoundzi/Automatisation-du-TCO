"""
persistence.py — Sauvegarde et chargement des projets (format JSON sécurisé).

Remplace pickle par JSON + gzip pour éliminer le risque d'exécution de
code arbitraire. Les DataFrames sont sérialisés via .to_dict(orient="records").

Format v3 : architecture multi-lots (1 projet = N lots isolés).
Format v2 : rétrocompatibilité — migré automatiquement vers v3 à la lecture.
"""

import gzip
import json
import os
import re as _re
import uuid
from decimal import Decimal

import pandas as pd

from config import PROJECTS_DIR, TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)

# Extension des fichiers projet (nouvelle version)
PROJECT_EXT = ".tco.json.gz"
# Ancienne extension (pickle) pour migration
LEGACY_EXT = ".tco"


def _validate_project_name(name: str) -> bool:
    """Vérifie que le nom de projet ne contient pas de séquences de path traversal."""
    if not name:
        return False
    # Rejeter toute tentative de traversal : .., /, \, :
    if ".." in name or "/" in name or "\\" in name or ":" in name:
        log.warning("Nom de projet suspect (path traversal) rejeté : %r", name)
        return False
    return True


def _project_path(name: str) -> str:
    """Retourne le chemin complet du fichier projet."""
    return os.path.join(PROJECTS_DIR, f"{name}{PROJECT_EXT}")


def _legacy_path(name: str) -> str:
    """Retourne le chemin de l'ancien format pickle."""
    return os.path.join(PROJECTS_DIR, f"{name}{LEGACY_EXT}")


# ---------------------------------------------------------------------------
# Migration v2 → v3
# ---------------------------------------------------------------------------


def _lot_stub_from_v2(data: dict) -> dict:
    """Transforme les données v2 plates en structure lot v3.

    Les champs tco_df / merged_df sont déjà des list[dict] (JSON désérialisé),
    on les passe directement sans appeler .to_dict().
    """
    meta = data.get("tco_meta") or {}
    lot_label_raw = ((meta.get("project_info") or {}).get("lot") or "").strip()
    lot_label = lot_label_raw or "LOT INCONNU"
    m = _re.search(r"\b(\d{2})\b", lot_label)
    return {
        "lot_id": uuid.uuid4().hex,
        "lot_label": lot_label,
        "lot_num": m.group(1) if m else "00",
        "tco_df": data.get("tco_df"),        # list[dict] ou None
        "tco_meta": meta,
        "tva_rate": data.get("tva_rate", TVA_DEFAULT),
        "merged_df": data.get("merged_df"),  # list[dict] ou None
        "all_alerts": data.get("all_alerts", []),
        "companies": data.get("company_data", {}),
    }


def _migrate_v2_to_v3(data: dict) -> dict:
    """Enveloppe un projet v2 dans la structure v3 (lot unique)."""
    log.info("Migration v2→v3 du projet '%s'", data.get("project_name", "?"))
    lot = _lot_stub_from_v2(data)
    return {
        "version": 3,
        "project_id": uuid.uuid4().hex,
        "project_name": data.get("project_name", ""),
        "created_at": "",
        "lots": [lot],
        "step": data.get("step", 0),
        "active_lot_id": lot["lot_id"],
    }


# ---------------------------------------------------------------------------
# Sérialisation helpers
# ---------------------------------------------------------------------------


def _json_default(obj):
    """Sérialise Decimal en float pour JSON."""
    if isinstance(obj, Decimal):
        return float(obj)
    return str(obj)


def _serialize_lot(lot: dict) -> dict:
    """Sérialise un lot (convertit les DataFrames en list[dict])."""
    lot_ser = dict(lot)
    for df_key in ("tco_df", "merged_df"):
        val = lot_ser.get(df_key)
        if hasattr(val, "to_dict"):
            lot_ser[df_key] = val.to_dict(orient="records")
    companies_ser = {}
    for comp_name, comp_info in lot.get("companies", {}).items():
        dpgf = comp_info["dpgf_df"]
        companies_ser[comp_name] = {
            "dpgf_df": dpgf.to_dict(orient="records") if hasattr(dpgf, "to_dict") else dpgf,
            "parse_alerts": comp_info["parse_alerts"],
            "filename": comp_info["filename"],
        }
    lot_ser["companies"] = companies_ser
    return lot_ser


def _deserialize_lot(lot_raw: dict) -> dict:
    """Désérialise un lot (convertit list[dict] en DataFrames)."""
    lot = dict(lot_raw)
    lot["tco_df"] = pd.DataFrame(lot["tco_df"]) if lot.get("tco_df") else None
    lot["merged_df"] = pd.DataFrame(lot["merged_df"]) if lot.get("merged_df") else None
    companies = {}
    for comp_name, comp_info in lot.get("companies", {}).items():
        companies[comp_name] = {
            "dpgf_df": pd.DataFrame(comp_info["dpgf_df"]),
            "parse_alerts": comp_info["parse_alerts"],
            "filename": comp_info["filename"],
        }
    lot["companies"] = companies
    return lot


# ---------------------------------------------------------------------------
# API publique
# ---------------------------------------------------------------------------


def save_project(name: str, session_state) -> tuple[bool, str]:
    """
    Sauvegarde l'état actuel dans un fichier JSON compressé (format v3).

    Args:
        name: nom du projet
        session_state: st.session_state contenant active_project et active_lot_id

    Returns:
        (success, message)
    """
    if not _validate_project_name(name):
        return False, "Nom de projet invalide (caractères interdits)."

    os.makedirs(PROJECTS_DIR, exist_ok=True)
    path = _project_path(name)

    active_project = session_state.get("active_project") or {}
    lots_serialized = [_serialize_lot(lot) for lot in active_project.get("lots", [])]

    data = {
        "version": 3,
        "project_id": active_project.get("project_id", ""),
        "project_name": name,
        "created_at": active_project.get("created_at", ""),
        "lots": lots_serialized,
        "step": session_state.get("step", 0),
        "active_lot_id": session_state.get("active_lot_id"),
    }

    try:
        with gzip.open(path, "wt", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, default=_json_default)
        log.info("Projet sauvegardé (JSON v3) : %s", name)

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

    Détecte automatiquement la version et migre v2→v3 si nécessaire.

    Args:
        name: nom du projet
        session_state: st.session_state à restaurer

    Returns:
        (success, message)
    """
    if not _validate_project_name(name):
        return False, "Nom de projet invalide (caractères interdits)."

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

        # Migration automatique v2 → v3
        version = data.get("version", 1)
        if version < 3:
            data = _migrate_v2_to_v3(data)

        lots = [_deserialize_lot(lot_raw) for lot_raw in data.get("lots", [])]

        session_state.active_project = {
            "project_id": data.get("project_id", ""),
            "project_name": name,
            "created_at": data.get("created_at", ""),
            "lots": lots,
        }
        session_state.active_lot_id = data.get("active_lot_id")
        session_state.step = data.get("step", 0)
        session_state.current_project = name

        log.info("Projet chargé (JSON v%d→v3) : %s, %d lot(s)", version, name, len(lots))
        return True, f"Projet '{name}' chargé ({len(lots)} lot(s))."
    except Exception as e:
        log.error("Erreur chargement projet %s : %s", name, e)
        return False, f"Erreur de lecture : {e}"


def _migrate_legacy_project(name: str, legacy_path: str, session_state) -> tuple[bool, str]:
    """
    L'ancien format pickle (.tco) n'est plus supporté pour des raisons de sécurité.

    Le format pickle permet l'exécution de code arbitraire lors du chargement.
    Les projets au format legacy doivent être recréés manuellement.
    Voir : https://docs.python.org/3/library/pickle.html#pickle-security
    """
    log.error("Chargement refusé — format pickle non sécurisé (SEC-PICKLE) : %s", legacy_path)
    # Proposer de supprimer le fichier dangereux
    try:
        os.remove(legacy_path)
        log.warning("Fichier pickle supprimé automatiquement : %s", legacy_path)
    except OSError as e:
        log.warning("Impossible de supprimer le fichier pickle %s : %s", legacy_path, e)

    return (
        False,
        f"Le projet '{name}' utilise un ancien format non sécurisé (pickle). "
        "Il a été supprimé automatiquement. Veuillez recréer ce projet.",
    )


def list_projects() -> list[str]:
    """Liste les noms de projets disponibles (JSON et legacy)."""
    if not os.path.exists(PROJECTS_DIR):
        return []

    names: set[str] = set()

    for f in os.listdir(PROJECTS_DIR):
        if f.endswith(PROJECT_EXT):
            # Retirer l'extension composée .tco.json.gz
            name = f.removesuffix(PROJECT_EXT)
            if _validate_project_name(name):
                names.add(name)
        elif f.endswith(LEGACY_EXT):
            # Ancien format pickle — signalé mais non chargeable
            name = os.path.splitext(f)[0]
            if _validate_project_name(name):
                names.add(name)

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
