"""
parser_tco.py — Lecture et validation du TCO modèle (fichier DPGF LOT).

Détecte dynamiquement la ligne d'en-tête, extrait les colonnes
Code/Désignation/Qu./U./Px U./Px tot. et les métadonnées du projet.
Classifie chaque ligne selon son type (section, recap, article, total, etc.)
en se basant sur la colonne Entete (col M).
"""

from __future__ import annotations

from decimal import Decimal

import pandas as pd

from core.utils import classify_row, find_column_index, open_excel_file
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _extract_project_info(df: pd.DataFrame, header_row_idx: int) -> dict[str, str]:
    """
    Extrait les informations du projet situées au-dessus de la ligne d'en-tête
    en cherchant des mots-clés.
    """
    info = {}
    keywords = {
        "region": ["région", "region"],
        "projet": ["projet", "opération", "operation"],
        "adresse": ["adresse", "lieu"],
        "phase": ["phase", "étape"],
        "lot": ["lot"],
    }

    # On parcourt les lignes au-dessus du header
    for r in range(header_row_idx):
        for c in range(min(8, df.shape[1])):  # On regarde les 8 premières colonnes
            val = str(df.iloc[r, c]).strip().lower()
            if not val:
                continue

            for key, kw_list in keywords.items():
                if key not in info and any(kw in val for kw in kw_list):
                    # La valeur est soit dans la même cellule après le label,
                    # soit dans la cellule suivante
                    if ":" in val:
                        info[key] = val.split(":", 1)[1].strip().upper()
                    elif c + 1 < df.shape[1]:
                        next_val = str(df.iloc[r, c + 1]).strip()
                        if next_val and next_val.lower() != "nan":
                            info[key] = next_val.upper()
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

    try:
        # open_excel_file : détecte engine, feuille et en-tête en un seul appel
        # (2 lectures au lieu de 3 — le probe est réutilisé comme df_raw)
        xl_file, sheet_name, df_raw, header_row_idx, _engine_kwargs = open_excel_file(filepath)

        project_info = _extract_project_info(df_raw, header_row_idx)

        # Lecture finale avec skiprows=header_row_idx (1 seule lecture supplémentaire)
        df_data = xl_file.parse(
            sheet_name,
            skiprows=header_row_idx,
            dtype=object,  # preserve codes comme strings
        )
        log.debug("En-tête trouvée index %d | projet=%s", header_row_idx, project_info)
    except Exception as e:
        log.error("Erreur de structure TCO: %s", e)
        return pd.DataFrame(), {
            "project_info": {},
            "header_row": 0,
            "sheet_name": "TCO",
            "filepath": filepath,
            "error": str(e),
        }

    def to_decimal(val: object) -> Decimal:
        if pd.isna(val) or not str(val).strip():
            return Decimal("0.0")
        try:
            return Decimal(str(val))
        except Exception:
            return Decimal("0.0")

    rows = []
    current_section_code = ""

    # Dynamic column mapping
    idx_code = find_column_index(df_data, ["code"], 0)
    idx_desig = find_column_index(df_data, ["désignation", "designation", "libellé"], 1)
    idx_qu = find_column_index(df_data, ["qu.", "quantité", "qte", "qté"], 2)
    idx_u = find_column_index(df_data, ["u", "unité"], 3)
    idx_pu = find_column_index(df_data, ["px u", "p.u", "prix u"], 4)
    idx_tot = find_column_index(df_data, ["px tot", "total ht", "prix tot"], 5)
    idx_entete = find_column_index(
        df_data, ["entete", "entête"]
    )  # None → COL_NOT_FOUND (-1) si absent

    for idx_in_df, xl_row in df_data.iterrows():
        row_idx = idx_in_df + header_row_idx + 2  # conversion en 1-indexed Excel row

        if len(xl_row) <= max(idx_code, idx_desig, idx_qu, idx_pu, idx_tot):
            continue

        code_raw = xl_row.iloc[idx_code]
        desig_raw = xl_row.iloc[idx_desig]
        qu_raw = xl_row.iloc[idx_qu]
        u = xl_row.iloc[idx_u]
        px_u_raw = xl_row.iloc[idx_pu]
        px_tot_raw = xl_row.iloc[idx_tot]
        entete = xl_row.iloc[idx_entete] if (idx_entete >= 0 and len(xl_row) > idx_entete) else None

        code_str = str(code_raw).strip() if pd.notna(code_raw) else ""
        desig_str = str(desig_raw).strip() if pd.notna(desig_raw) else ""
        ent_str = str(entete).strip() if pd.notna(entete) else ""

        qu = to_decimal(qu_raw)
        px_u = to_decimal(px_u_raw)
        px_tot = to_decimal(px_tot_raw)

        # has_price : Qu et PU non nuls indiquent un article (fallback quand Entete absent)
        has_price_tco = qu > Decimal("0.0") and px_u > Decimal("0.0")
        row_type = classify_row(code_str, desig_str, ent_str, has_price=has_price_tco)

        # Correction : les codes courts (≤2 segments, ex "02.2") restent section_header
        # même s'ils portent un prix forfaitaire direct (Qu.=1, Px_U_HT=X).
        # Sans ce correctif, classify_row les retourne "article" via has_price,
        # ce qui crée un doublon visuel (article + recap affichent le même montant).
        if row_type == "article" and code_str:
            _segs = [p for p in code_str.split(".") if p.strip()]
            if len(_segs) <= 2:
                row_type = "section_header"

        # Lignes non classifiables (titres de document, textes libres) → ignorées
        # Exception : si la désignation contient des mots-clés financiers (montant, tva, ttc)
        # la ligne est re-classifiée en total_line pour ne pas la supprimer.
        if row_type == "other":
            _d = desig_str.lower()
            if "montant" in _d or "tva" in _d or "ttc" in _d:
                row_type = "total_line"
            else:
                continue

        if row_type == "section_header":
            current_section_code = code_str

        parent_code = current_section_code if row_type == "recap" else ""

        rows.append(
            {
                "Code": code_str,
                "Désignation": desig_str,
                "Qu.": qu,
                "U": str(u).strip() if pd.notna(u) else "",
                "Px_U_HT": px_u,
                "Px_Tot_HT": px_tot,
                "Entete": ent_str,
                "row_type": row_type,
                "original_row": row_idx,
                "parent_code": parent_code,
            }
        )

    tco_df = pd.DataFrame(rows)
    log.info(
        "TCO parsé : %d lignes (%d articles, %d sections)",
        len(tco_df),
        len(tco_df[tco_df["row_type"] == "article"]),
        len(tco_df[tco_df["row_type"] == "section_header"]),
    )

    meta = {
        "project_info": project_info,
        "header_row": header_row_idx + 1,
        "sheet_name": "TCO",
        "filepath": filepath,
    }

    return tco_df, meta
