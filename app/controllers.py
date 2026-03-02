"""
controllers.py — Orchestration métier sans dépendance Streamlit (A5).

Règle : zéro import streamlit. Toutes les fonctions prennent leurs
données en paramètre et retournent des valeurs — pas d'effets de bord
sur session_state. Cela rend la logique métier testable unitairement.
"""

from __future__ import annotations

import re

from config import COMPANY_NAME_MAX_LEN
from core.merger import merge_all_companies, merge_company_into_tco

# APP-1 : & et ' retirés — ils peuvent servir à s'échapper de contextes HTML/SQL
COMPANY_PATTERN = re.compile(r"^[A-Za-zÀ-ÖØ-öø-ÿ0-9 \-_.()]+$")

# Regex pré-compilées pour _normalize_filename()
_RE_NON_ALPHANUM = re.compile(r"[^A-Z0-9]")
_RE_MULTI_UNDERSCORE = re.compile(r"_+")


def validate_company_name(name: str) -> tuple[str | None, str | None]:
    """Valide et normalise un nom d'entreprise.

    Args:
        name: nom brut saisi par l'utilisateur.

    Returns:
        (name_clean, None) si valide, (None, error_msg) sinon.
    """
    name = name.strip()
    if not name:
        return None, "Le nom ne peut pas être vide."
    if len(name) > COMPANY_NAME_MAX_LEN:
        return None, f"Nom trop long (max {COMPANY_NAME_MAX_LEN} caractères)."
    if not COMPANY_PATTERN.match(name):
        return None, "Nom invalide (caractères spéciaux interdits)."
    return name, None


def normalize_filename(s: str) -> str:
    """Normalise une chaîne pour l'inclure dans un nom de fichier Excel.

    Convertit en majuscules, remplace tout caractère non alphanumérique par
    un underscore, puis compresse les underscores consécutifs.

    Args:
        s: chaîne brute (nom de projet ou de lot).

    Returns:
        Chaîne normalisée sans underscore de début/fin.

    Examples:
        >>> normalize_filename("Chantier Bordeaux 2026")
        'CHANTIER_BORDEAUX_2026'
    """
    norm = _RE_NON_ALPHANUM.sub("_", s.upper())
    return _RE_MULTI_UNDERSCORE.sub("_", norm).strip("_")


def rebuild_merged_tco(
    tco_df,
    company_data: dict,
    tva_rate: float,
    merged_df=None,
    new_companies: list[str] | None = None,
) -> tuple:
    """Reconstruit ou met à jour le TCO fusionné (sans effets de bord).

    Args:
        tco_df: DataFrame de base (DPGF vierge du lot).
        company_data: dict {nom_entreprise: {dpgf_df, parse_alerts, filename}}.
        tva_rate: taux de TVA applicable (ex. 0.20).
        merged_df: DataFrame actuellement fusionné (pour fusion incrémentale).
        new_companies: si fourni ET merged_df présent → fusion incrémentale
            (ajoute uniquement ces entreprises). Sinon reconstruction complète.

    Returns:
        Tuple (merged_df, alerts_list) :
            - merged_df : DataFrame fusionné mis à jour.
            - alerts_list : nouvelles alertes uniquement (chemin incrémental)
              ou toutes les alertes (reconstruction complète).
    """
    if new_companies and merged_df is not None:
        # Fusion incrémentale : O(k) avec k = len(new_companies)
        merged = merged_df.copy()
        new_alerts: list[dict] = []
        for comp_name in new_companies:
            comp = company_data[comp_name]
            merged, merge_alerts = merge_company_into_tco(
                merged, comp["dpgf_df"], comp_name, tva_rate=tva_rate
            )
            for alert in comp.get("parse_alerts", []):
                alert["company"] = comp_name
            for alert in merge_alerts:
                alert["company"] = comp_name
            new_alerts.extend(comp.get("parse_alerts", []))
            new_alerts.extend(merge_alerts)
        return merged, new_alerts

    # Reconstruction complète depuis le TCO de base
    return merge_all_companies(tco_df, company_data, tva_rate=tva_rate)
