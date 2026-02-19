"""
parser_tco.py — Lecture et validation du TCO modèle (fichier DPGF LOT).

Détecte dynamiquement la ligne d'en-tête, extrait les colonnes
Code/Désignation/Qu./U./Px U./Px tot. et les métadonnées du projet.
Classifie chaque ligne selon son type (section, recap, article, total, etc.)
en se basant sur la colonne Entete (col M).
"""

import openpyxl
import pandas as pd

from core.utils import find_header_row, classify_row
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _extract_project_info(ws, header_row):
    """
    Extrait les informations du projet situées au-dessus de la ligne d'en-tête.
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


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_tco(filepath):
    """
    Lit un fichier TCO modèle (DPGF LOT .xlsx).

    Returns:
        tco_df (DataFrame) : colonnes
            [Code, Désignation, Qu., U, Px_U_HT, Px_Tot_HT,
             Entete, row_type, original_row, parent_code]
        meta (dict) : project_info, header_row, sheet_name, filepath
    """
    log.info("Lecture TCO : %s", filepath)

    wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
    ws = wb.active
    sheet_name = ws.title

    header_row   = find_header_row(ws)
    project_info = _extract_project_info(ws, header_row)
    log.debug("En-tête trouvée ligne %d | projet=%s", header_row, project_info)

    rows = []
    current_section_code = ""

    for row_idx, xl_row in enumerate(
        ws.iter_rows(min_row=header_row + 1, values_only=True),
        start=header_row + 1,
    ):
        if len(xl_row) < 6:
            continue

        code_raw  = xl_row[0]
        desig_raw = xl_row[1]
        qu        = xl_row[2]
        u         = xl_row[3]
        px_u      = xl_row[4]
        px_tot    = xl_row[5]
        entete    = xl_row[12] if len(xl_row) > 12 else None

        code_str  = str(code_raw).strip()  if code_raw  else ""
        desig_str = str(desig_raw).strip() if desig_raw else ""
        ent_str   = str(entete).strip()    if entete    else ""

        row_type = classify_row(code_str, desig_str, ent_str)

        if row_type == "section_header":
            current_section_code = code_str

        parent_code = current_section_code if row_type == "recap" else ""

        rows.append({
            "Code":         code_str,
            "Désignation":  desig_str,
            "Qu.":          qu,
            "U":            str(u).strip() if u else "",
            "Px_U_HT":      px_u,
            "Px_Tot_HT":    px_tot,
            "Entete":       ent_str,
            "row_type":     row_type,
            "original_row": row_idx,
            "parent_code":  parent_code,
        })

    wb.close()
    tco_df = pd.DataFrame(rows)
    log.info(
        "TCO parsé : %d lignes (%d articles, %d sections)",
        len(tco_df),
        len(tco_df[tco_df["row_type"] == "article"]),
        len(tco_df[tco_df["row_type"] == "section_header"]),
    )

    meta = {
        "project_info": project_info,
        "header_row":   header_row,
        "sheet_name":   sheet_name,
        "filepath":     filepath,
    }

    return tco_df, meta
