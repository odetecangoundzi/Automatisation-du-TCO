"""
parser_tco.py — Lecture et validation du TCO modèle (fichier DPGF LOT).

Détecte dynamiquement la ligne d'en-tête, extrait les colonnes
Code/Désignation/Qu./U./Px U./Px tot. et les métadonnées du projet.
"""

import pandas as pd
import openpyxl
import re


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_header_row(ws):
    """
    Parcourt les lignes pour trouver celle contenant 'Code' en col A
    et 'Désignation' en col B. Retourne le numéro de ligne (1-indexed).
    """
    for row_idx in range(1, min(ws.max_row + 1, 20)):
        a_val = ws.cell(row=row_idx, column=1).value
        b_val = ws.cell(row=row_idx, column=2).value
        if (
            a_val and str(a_val).strip().lower() == "code"
            and b_val and "signation" in str(b_val).strip().lower()
        ):
            return row_idx
    raise ValueError("Impossible de trouver la ligne d'en-tête (Code | Désignation).")


def _extract_project_info(ws, header_row):
    """
    Extrait les informations du projet situées au-dessus de la ligne d'en-tête.
    Retourne un dict avec les clés : region, projet, adresse, phase, lot.
    """
    info = {}
    labels = ["region", "projet", "adresse", "phase", "lot"]
    idx = 0
    for row_idx in range(1, header_row):
        val = ws.cell(row=row_idx, column=1).value
        if val and str(val).strip():
            if idx < len(labels):
                info[labels[idx]] = str(val).strip()
                idx += 1
    return info


def _is_total_row(code_val):
    """Vérifie si la ligne est une ligne de total."""
    if not code_val:
        return False
    return str(code_val).strip().lower().startswith("total")


def _is_summary_row(code_val):
    """Vérifie si la ligne est un résumé (Montant HT, Montant TTC, etc.)."""
    if not code_val:
        return False
    s = str(code_val).strip().lower()
    return any(kw in s for kw in ["montant", "tva", "taux"])


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_tco(filepath):
    """
    Lit un fichier TCO modèle (DPGF LOT .xlsx).

    Retourne :
        tco_df : DataFrame avec colonnes
            [Code, Désignation, Qu., U, Px_U_HT, Px_Tot_HT, Entete, row_type]
        meta   : dict contenant project_info, header_row, sheet_name
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active
    sheet_name = ws.title

    header_row = _find_header_row(ws)
    project_info = _extract_project_info(ws, header_row)

    rows = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=row_idx, column=1).value
        designation = ws.cell(row=row_idx, column=2).value
        qu = ws.cell(row=row_idx, column=3).value
        u = ws.cell(row=row_idx, column=4).value
        px_u = ws.cell(row=row_idx, column=5).value
        px_tot = ws.cell(row=row_idx, column=6).value
        entete = ws.cell(row=row_idx, column=13).value  # col M

        # Déterminer le type de ligne
        if _is_total_row(code):
            row_type = "total"
        elif _is_summary_row(code) or _is_summary_row(designation):
            row_type = "summary"
        elif entete and "Recap" in str(entete):
            row_type = "recap"
        elif code and designation:
            row_type = "data"
        elif not code and not designation:
            row_type = "empty"
        else:
            row_type = "other"

        rows.append({
            "Code": str(code).strip() if code else "",
            "Désignation": str(designation).strip() if designation else "",
            "Qu.": qu,
            "U": str(u).strip() if u else "",
            "Px_U_HT": px_u,
            "Px_Tot_HT": px_tot,
            "Entete": str(entete).strip() if entete else "",
            "row_type": row_type,
            "original_row": row_idx,
        })

    tco_df = pd.DataFrame(rows)

    meta = {
        "project_info": project_info,
        "header_row": header_row,
        "sheet_name": sheet_name,
        "filepath": filepath,
    }

    wb.close()
    return tco_df, meta
