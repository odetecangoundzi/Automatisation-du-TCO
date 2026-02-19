"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Gère les lignes de total et les sous-totaux.
"""

import pandas as pd
import re


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_code(code):
    """Normalise un code pour la comparaison (strip, lowercase)."""
    if not code:
        return ""
    return str(code).strip()


def _is_section_code(code):
    """
    Vérifie si le code est un code de section (ex: 01.2, 01.2.1)
    par opposition à un article (ex: 01.2.1.1.1).
    """
    if not code:
        return False
    parts = str(code).strip().split(".")
    return len(parts) <= 3


# ---------------------------------------------------------------------------
# Main merger
# ---------------------------------------------------------------------------

def merge_company_into_tco(tco_df, dpgf_df, company_name):
    """
    Fusionne un DPGF normalisé dans le TCO.

    Args:
        tco_df      : DataFrame du TCO modèle (de parse_tco)
        dpgf_df     : DataFrame normalisé du DPGF (de parse_dpgf)
        company_name: nom de l'entreprise (ex: "MAB SUD-OUEST")

    Retourne :
        merged_df : DataFrame avec les colonnes entreprise ajoutées
        alerts    : liste d'alertes (codes non trouvés)
    """
    merged_df = tco_df.copy()
    alerts = []

    # Noms des colonnes entreprise
    col_qu = f"{company_name}_Qu."
    col_pu = f"{company_name}_Px_U_HT"
    col_tot = f"{company_name}_Px_Tot_HT"
    col_com = f"{company_name}_Commentaire"

    # Initialiser les colonnes entreprise
    merged_df[col_qu] = None
    merged_df[col_pu] = None
    merged_df[col_tot] = None
    merged_df[col_com] = None

    # Index des codes TCO pour lookup rapide
    tco_code_index = {}
    for idx, row in merged_df.iterrows():
        code = _normalize_code(row["Code"])
        if code and row["row_type"] not in ("total", "empty"):
            tco_code_index[code] = idx

    # Fusion par code
    dpgf_data_rows = dpgf_df[dpgf_df["row_type"] == "data"]
    matched_codes = set()

    for _, dpgf_row in dpgf_data_rows.iterrows():
        code = _normalize_code(dpgf_row["Code"])
        if not code:
            continue

        if code in tco_code_index:
            idx = tco_code_index[code]
            merged_df.at[idx, col_qu] = dpgf_row["Qu."]
            merged_df.at[idx, col_pu] = dpgf_row["Px_U_HT"]
            merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
            merged_df.at[idx, col_com] = dpgf_row.get("Commentaire", "")
            matched_codes.add(code)
        else:
            alerts.append({
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Code '{code}' du DPGF non trouvé dans le TCO",
            })

    # Calculer les sous-totaux pour les lignes de section et de total
    _compute_subtotals(merged_df, col_tot)

    return merged_df, alerts


def _compute_subtotals(df, total_col):
    """
    Recalcule les sous-totaux des lignes 'total' et des sections
    en additionnant les Px_Tot_HT des lignes enfants.
    """
    # Identifier les lignes de total et leur associer les enfants
    for idx, row in df.iterrows():
        if row["row_type"] == "total":
            # Le total porte le nom de la section, trouver les enfants
            # en remontant jusqu'à trouver la section parente
            total_sum = 0.0
            found_parent = False

            # Parcourir les lignes au-dessus du total pour sommer les enfants
            for prev_idx in range(idx - 1, -1, -1):
                prev_row = df.iloc[prev_idx]
                prev_code = _normalize_code(prev_row["Code"])

                # Arrêt quand on trouve un autre total ou le début
                if prev_row["row_type"] == "total":
                    break

                # Sommer les valeurs des lignes de données
                val = prev_row.get(total_col)
                if val is not None and prev_row["row_type"] == "data":
                    try:
                        total_sum += float(val)
                    except (ValueError, TypeError):
                        pass

            if total_sum > 0:
                df.at[idx, total_col] = total_sum

    # Calculer les totaux de section (ex: 01.2 somme les enfants 01.2.x.x.x)
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or row["row_type"] in ("total", "empty", "recap"):
            continue

        if _is_section_code(code):
            section_total = 0.0
            # Chercher les enfants directs de cette section
            for child_idx, child_row in df.iterrows():
                child_code = _normalize_code(child_row["Code"])
                if (
                    child_code
                    and child_code != code
                    and child_code.startswith(code + ".")
                    and child_row["row_type"] == "data"
                ):
                    val = child_row.get(total_col)
                    if val is not None:
                        try:
                            section_total += float(val)
                        except (ValueError, TypeError):
                            pass

            if section_total > 0:
                df.at[idx, total_col] = section_total
