"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Calcule les sous-totaux en se basant sur la hiérarchie Code (01.X → 01.X.Y...).
Les lignes "recap" (Code vide, Entete Bord_xxx_Recap) reçoivent le total
de leur section parente.
"""

import pandas as pd
from config import TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_code(code):
    """Normalise un code pour la comparaison."""
    return str(code).strip() if code else ""


def _build_section_index(df):
    """
    PERF-1 : Pré-indexe les section_headers par code pour lookup O(1).
    Retourne un dict {code: dataframe_index}.
    """
    return {
        _normalize_code(row["Code"]): idx
        for idx, row in df.iterrows()
        if row["row_type"] == "section_header" and row["Code"]
    }


def _get_children_total(df, parent_code, total_col, children_index):
    """
    Calcule la somme des valeurs d'une colonne pour tous les enfants
    (articles ET sub_sections) d'une section.

    Args:
        children_index : dict {code: [idx, ...]} pré-calculé
    """
    prefix = parent_code + "."
    total  = 0.0
    count  = 0
    for code, idx_list in children_index.items():
        if not code.startswith(prefix):
            continue
        for idx in idx_list:
            row = df.loc[idx]
            if row["row_type"] in ("article", "sub_section"):
                val = row.get(total_col)
                if val is not None:
                    try:
                        total += float(val)
                        count += 1
                    except (ValueError, TypeError):
                        pass
    return total


# ---------------------------------------------------------------------------
# Main merger
# ---------------------------------------------------------------------------

def merge_company_into_tco(tco_df, dpgf_df, company_name):
    """
    Fusionne un DPGF normalisé dans le TCO.

    Args:
        tco_df       : DataFrame du TCO modèle (de parse_tco)
        dpgf_df      : DataFrame normalisé du DPGF (de parse_dpgf)
        company_name : nom de l'entreprise

    Returns:
        merged_df : DataFrame avec colonnes entreprise ajoutées
        alerts    : liste d'alertes (codes non trouvés)
    """
    log.info("Fusion DPGF pour l'entreprise : %s", company_name)
    merged_df = tco_df.copy()
    alerts    = []

    col_qu  = f"{company_name}_Qu."
    col_pu  = f"{company_name}_Px_U_HT"
    col_tot = f"{company_name}_Px_Tot_HT"
    col_com = f"{company_name}_Commentaire"

    merged_df[col_qu]  = None
    merged_df[col_pu]  = None
    merged_df[col_tot] = None
    merged_df[col_com] = None

    # PERF-1 : index TCO codes → O(1) lookup
    tco_code_index = {}
    for idx, row in merged_df.iterrows():
        code = _normalize_code(row["Code"])
        if code and row["row_type"] not in ("empty", "recap", "recap_summary"):
            if code not in tco_code_index:
                tco_code_index[code] = idx

    # Fusion articles + sub_sections du DPGF
    dpgf_data = dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])]
    matched_count = 0

    for _, dpgf_row in dpgf_data.iterrows():
        code = _normalize_code(dpgf_row["Code"])
        if not code:
            continue
        if code in tco_code_index:
            idx = tco_code_index[code]
            merged_df.at[idx, col_qu]  = dpgf_row["Qu."]
            merged_df.at[idx, col_pu]  = dpgf_row["Px_U_HT"]
            merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
            merged_df.at[idx, col_com] = dpgf_row.get("Commentaire", "")
            matched_count += 1
        else:
            alerts.append({
                "type":    "warning",
                "color":   "orange",
                "code":    code,
                "message": f"Code '{code}' du DPGF non trouvé dans le TCO",
            })

    log.info(
        "Fusion terminée : %d lignes matchées, %d non trouvées",
        matched_count, len(alerts)
    )

    _compute_section_totals(merged_df, col_tot)
    return merged_df, alerts


# ---------------------------------------------------------------------------
# Calcul des sous-totaux
# ---------------------------------------------------------------------------

def _compute_section_totals(df, total_col, tva_rate=TVA_DEFAULT):
    """
    Recalcule les totaux pour :
      1. section_header (01.X) : somme directe des articles + sub_sections
      2. recap (Code vide)     : reçoit le total de sa section parente
      3. recap_summary         : reçoit le total de leur section
      4. total_line            : Montant HT / TVA / TTC

    Args:
        total_col : colonne à calculer (ex: "MAB SUD-OUEST_Px_Tot_HT")
        tva_rate  : taux de TVA (paramétrable)
    """
    children_index: dict[str, list] = {}
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if code:
            children_index.setdefault(code, []).append(idx)

    section_header_index = _build_section_index(df)

    # Passe 1 : calculer les totaux des section_headers
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or row["row_type"] != "section_header":
            continue
        total = _get_children_total(df, code, total_col, children_index)
        if total > 0:
            df.at[idx, total_col] = total

    # Passe 2 : propager vers les lignes recap
    for idx, row in df.iterrows():
        if row["row_type"] != "recap":
            continue
        parent = _normalize_code(row.get("parent_code", ""))
        if parent and parent in section_header_index:
            s_idx = section_header_index[parent]
            val   = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val

    # Passe 3 : recap_summary
    for idx, row in df.iterrows():
        if row["row_type"] != "recap_summary":
            continue
        code = _normalize_code(row["Code"])
        if code and code in section_header_index:
            s_idx = section_header_index[code]
            val   = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val

    # Passe 4 : Montant HT / TVA / TTC
    montant_ht = 0.0
    for idx, row in df.iterrows():
        if row["row_type"] == "section_header":
            val = row.get(total_col)
            if val is not None:
                try:
                    montant_ht += float(val)
                except (ValueError, TypeError):
                    pass

    if montant_ht > 0:
        tva         = montant_ht * tva_rate
        montant_ttc = montant_ht + tva
        term_map = {"montant ht": montant_ht, "tva": tva, "ttc": montant_ttc}
        
        for idx, row in df.iterrows():
            if row["row_type"] != "total_line":
                continue
            desig = str(row.get("Désignation", "")).strip().lower()
            for key, val in term_map.items():
                if key in desig:
                    df.at[idx, total_col] = val
                    break

    # Colonne de base Px_Tot_HT (si on vient de calculer une colonne entreprise)
    if total_col != "Px_Tot_HT":
        _compute_ht_tva_ttc_base(df, tva_rate)


def _compute_ht_tva_ttc_base(df, tva_rate=TVA_DEFAULT):
    """Calcule Montant HT, TVA, TTC pour la colonne de base Px_Tot_HT."""
    montant_ht = sum(
        float(row["Px_Tot_HT"])
        for _, row in df.iterrows()
        if row["row_type"] == "section_header"
        and row.get("Px_Tot_HT") is not None
        and isinstance(row["Px_Tot_HT"], (int, float))
    )
    if montant_ht <= 0:
        return

    tva         = montant_ht * tva_rate
    montant_ttc = montant_ht + tva
    term_map    = {"montant ht": montant_ht, "tva": tva, "ttc": montant_ttc}

    for idx, row in df.iterrows():
        if row["row_type"] != "total_line":
            continue
        desig = str(row.get("Désignation", "")).strip().lower()
        for key, val in term_map.items():
            if key in desig:
                df.at[idx, "Px_Tot_HT"] = val
                break
