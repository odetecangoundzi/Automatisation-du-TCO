"""
parser_tco.py — Lecture et validation du TCO modèle (fichier DPGF LOT).

Détecte dynamiquement la ligne d'en-tête, extrait les colonnes
Code/Désignation/Qu./U./Px U./Px tot. et les métadonnées du projet.
Classifie chaque ligne selon son type (section, recap, article, total, etc.)
en se basant sur la colonne Entete (col M).
"""

from __future__ import annotations

import os
import pandas as pd

from core.utils import find_header_row, classify_row
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _extract_project_info(df: pd.DataFrame, header_row_idx: int) -> dict[str, str]:
    """
    Extrait les informations du projet situées au-dessus de la ligne d'en-tête.
    """
    info = {}
    labels = ["region", "projet", "adresse", "phase", "lot"]
    label_idx = 0
    
    for r in range(header_row_idx):
        val = df.iloc[r, 0]
        if pd.notna(val) and str(val).strip():
            if label_idx < len(labels):
                info[labels[label_idx]] = str(val).strip()
                label_idx += 1
    return info


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_tco(filepath: str) -> tuple[pd.DataFrame, dict]:
    """
    Lit un fichier TCO modèle (DPGF LOT XLSX, XLS, XLSB).

    Returns:
        tco_df (DataFrame) : colonnes
            [Code, Désignation, Qu., U, Px_U_HT, Px_Tot_HT,
             Entete, row_type, original_row, parent_code]
        meta (dict) : project_info, header_row, sheet_name, filepath
    """
    log.info("Lecture TCO : %s", filepath)

    ext = os.path.splitext(filepath)[1].lower()
    engine = None
    if ext == ".xls": engine = "xlrd"
    elif ext == ".xlsb": engine = "pyxlsb"
    elif ext in [".xlsx", ".xlsm"]: engine = "openpyxl"

    try:
        # Lecture sans header pour trouver la table
        df_raw = pd.read_excel(filepath, engine=engine, header=None)
        header_row_idx = find_header_row(df_raw)
        
        project_info = _extract_project_info(df_raw, header_row_idx)
        
        # Re-lecture avec le bon header
        df_data = pd.read_excel(filepath, engine=engine, skiprows=header_row_idx)
        log.debug("En-tête trouvée index %d | projet=%s", header_row_idx, project_info)
    except Exception as e:
        log.error("Erreur de structure TCO: %s", e)
        return pd.DataFrame(), {
            "project_info": {}, "header_row": 0, "sheet_name": "TCO", "filepath": filepath, "error": str(e)
        }

    rows = []
    current_section_code = ""

    for idx_in_df, xl_row in df_data.iterrows():
        row_idx = idx_in_df + header_row_idx + 2 # conversion en 1-indexed Excel row
        
        if len(xl_row) < 6:
            continue

        code_raw  = xl_row.iloc[0]
        desig_raw = xl_row.iloc[1]
        qu        = xl_row.iloc[2]
        u         = xl_row.iloc[3]
        px_u      = xl_row.iloc[4]
        px_tot    = xl_row.iloc[5]
        entete    = xl_row.iloc[12] if len(xl_row) > 12 else None

        code_str  = str(code_raw).strip()  if pd.notna(code_raw)  else ""
        desig_str = str(desig_raw).strip() if pd.notna(desig_raw) else ""
        ent_str   = str(entete).strip()    if pd.notna(entete)    else ""

        row_type = classify_row(code_str, desig_str, ent_str)

        if row_type == "section_header":
            current_section_code = code_str

        parent_code = current_section_code if row_type == "recap" else ""

        rows.append({
            "Code":         code_str,
            "Désignation":  desig_str,
            "Qu.":          qu,
            "U":            str(u).strip() if pd.notna(u) else "",
            "Px_U_HT":      px_u,
            "Px_Tot_HT":    px_tot,
            "Entete":       ent_str,
            "row_type":     row_type,
            "original_row": row_idx,
            "parent_code":  parent_code,
        })

    tco_df = pd.DataFrame(rows)
    log.info(
        "TCO parsé : %d lignes (%d articles, %d sections)",
        len(tco_df),
        len(tco_df[tco_df["row_type"] == "article"]),
        len(tco_df[tco_df["row_type"] == "section_header"]),
    )

    meta = {
        "project_info": project_info,
        "header_row":   header_row_idx + 1,
        "sheet_name":   "TCO",
        "filepath":     filepath,
    }

    return tco_df, meta

