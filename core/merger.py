"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Calcule les sous-totaux en se basant sur la hiérarchie Code (01.X → 01.X.Y...).
Les lignes "recap" (Code vide, Entete Bord_xxx_Recap) reçoivent le total
de leur section parente.
"""

import pandas as pd


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_code(code):
    """Normalise un code pour la comparaison."""
    if not code:
        return ""
    return str(code).strip()


def _get_children_total(df, parent_code, total_col):
    """
    Calcule la somme des Px_Tot_HT d'une colonne donnée pour tous
    les enfants (articles ET sub_sections) d'une section.
    Les sub_sections dans ce modèle portent des valeurs indépendantes
    de leurs propres enfants, il faut donc les additionner.
    """
    prefix = parent_code + "."
    total = 0.0
    for _, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or not code.startswith(prefix):
            continue
        if row["row_type"] in ("article", "sub_section"):
            val = row.get(total_col)
            if val is not None:
                try:
                    total += float(val)
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
        tco_df      : DataFrame du TCO modèle (de parse_tco)
        dpgf_df     : DataFrame normalisé du DPGF (de parse_dpgf)
        company_name: nom de l'entreprise

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
        if code and row["row_type"] not in ("empty", "recap", "recap_summary"):
            # Stocker seulement la première occurrence si duplicata
            if code not in tco_code_index:
                tco_code_index[code] = idx

    # Fusion par code — articles et sub_sections du DPGF
    dpgf_data = dpgf_df[
        dpgf_df["row_type"].isin(["article", "sub_section"])
    ]

    for _, dpgf_row in dpgf_data.iterrows():
        code = _normalize_code(dpgf_row["Code"])
        if not code:
            continue

        if code in tco_code_index:
            idx = tco_code_index[code]
            merged_df.at[idx, col_qu] = dpgf_row["Qu."]
            merged_df.at[idx, col_pu] = dpgf_row["Px_U_HT"]
            merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
            merged_df.at[idx, col_com] = dpgf_row.get("Commentaire", "")
        else:
            alerts.append({
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Code '{code}' du DPGF non trouvé dans le TCO",
            })

    # --- Calcul des sous-totaux ---
    _compute_section_totals(merged_df, col_tot)

    return merged_df, alerts


def _compute_section_totals(df, total_col):
    """
    Recalcule les totaux pour :
      1. section_header (01.X) : somme directe des articles + sub_sections
      2. recap (Code vide)     : reçoit le total de sa section parente
      3. recap_summary         : reçoit le total de leur section

    Note : les sub_sections portent déjà leurs valeurs propres depuis
    la fusion DPGF, elles ne sont PAS recalculées.
    """
    # Passe 1 : calculer les totaux des section_headers
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code:
            continue

        if row["row_type"] == "section_header":
            total = _get_children_total(df, code, total_col)
            if total > 0:
                df.at[idx, total_col] = total

    # Passe 2 : propager vers les lignes recap (Code vide)
    for idx, row in df.iterrows():
        if row["row_type"] == "recap":
            parent = _normalize_code(row.get("parent_code", ""))
            if parent:
                for s_idx, s_row in df.iterrows():
                    if (
                        _normalize_code(s_row["Code"]) == parent
                        and s_row["row_type"] == "section_header"
                    ):
                        val = s_row.get(total_col)
                        if val is not None:
                            df.at[idx, total_col] = val
                        break

    # Passe 3 : recap_summary (table récap en bas)
    for idx, row in df.iterrows():
        if row["row_type"] == "recap_summary":
            code = _normalize_code(row["Code"])
            if code:
                for s_idx, s_row in df.iterrows():
                    if (
                        _normalize_code(s_row["Code"]) == code
                        and s_row["row_type"] == "section_header"
                    ):
                        val = s_row.get(total_col)
                        if val is not None:
                            df.at[idx, total_col] = val
                        break

    # Passe 4 : Montant HT / TVA / Montant TTC (lignes total_line)
    # Montant HT = somme de tous les section_headers
    montant_ht = 0.0
    for _, row in df.iterrows():
        if row["row_type"] == "section_header":
            val = row.get(total_col)
            if val is not None:
                try:
                    montant_ht += float(val)
                except (ValueError, TypeError):
                    pass

    tva_rate = 0.20
    tva = montant_ht * tva_rate
    montant_ttc = montant_ht + tva

    # Affecter aux lignes total_line selon leur Désignation
    for idx, row in df.iterrows():
        if row["row_type"] == "total_line":
            desig = str(row.get("Désignation", "")).strip().lower()
            if "montant ht" in desig:
                df.at[idx, total_col] = montant_ht
            elif "tva" in desig:
                df.at[idx, total_col] = tva
            elif "montant ttc" in desig or "ttc" in desig:
                df.at[idx, total_col] = montant_ttc

    # Aussi calculer pour la colonne base Px_Tot_HT (estimation)
    if total_col != "Px_Tot_HT":
        # Déjà fait pour la colonne entreprise, faire aussi pour la base
        _compute_ht_tva_ttc_base(df)


def _compute_ht_tva_ttc_base(df):
    """
    Calcule Montant HT, TVA, TTC pour la colonne de base Px_Tot_HT.
    """
    montant_ht = 0.0
    for _, row in df.iterrows():
        if row["row_type"] == "section_header":
            val = row.get("Px_Tot_HT")
            if val is not None:
                try:
                    montant_ht += float(val)
                except (ValueError, TypeError):
                    pass

    if montant_ht <= 0:
        return

    tva_rate = 0.20
    tva = montant_ht * tva_rate
    montant_ttc = montant_ht + tva

    for idx, row in df.iterrows():
        if row["row_type"] == "total_line":
            desig = str(row.get("Désignation", "")).strip().lower()
            if "montant ht" in desig:
                df.at[idx, "Px_Tot_HT"] = montant_ht
            elif "tva" in desig:
                df.at[idx, "Px_Tot_HT"] = tva
            elif "montant ttc" in desig or "ttc" in desig:
                df.at[idx, "Px_Tot_HT"] = montant_ttc

