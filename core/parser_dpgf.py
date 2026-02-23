"""
parser_dpgf.py — Normalisation et extraction des annotations des DPGF entreprise.

Gère les textes dans les cellules numériques (SANS OBJET, COMPRIS, nc, P-M),
extrait les annotations en colonne Commentaire, et détecte les erreurs.
Utilise la colonne Entete (col M) pour classifier chaque ligne.
"""

from __future__ import annotations
import os
import re
import pandas as pd

from core.utils import find_header_row, classify_row
from config import TOTAL_TOLERANCE_ABS, TOTAL_TOLERANCE_REL
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

KEYWORDS = {
    "sans objet": "so",
    "compris":    "compris",
    "nc":         "nc",
    "p-m":        "pm",
    "inclus":     "inclus",
    "néant":      "néant",
    "so":         "so",
    "pm":         "pm",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _clean_numeric(value: int | float | str | None) -> tuple[float, str]:
    """
    Nettoie une valeur potentiellement numérique.
    Retourne (nombre_float, texte_annotation).
    """
    if pd.isna(value):
        return 0.0, ""
    if isinstance(value, (int, float)):
        return float(value), ""

    text = str(value).strip()
    if not text:
        return 0.0, ""

    text_lower = text.lower().strip()
    for keyword, abbrev in KEYWORDS.items():
        if keyword in text_lower:
            return 0.0, abbrev

    cleaned = text.replace(" ", "").replace("\u00a0", "").replace(",", ".")
    match   = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if match:
        try:
            number    = float(match.group())
            remaining = re.sub(r"-?\d+(?:\.\d+)?", "", cleaned).strip()
            remaining = remaining.strip("()[]{}/ ")
            return number, remaining
        except ValueError:
            pass

    return 0.0, text


def _check_total_coherence(
    qu_val: float, pu_val: float, total_val: float, row_idx: int, code: str
) -> dict | None:
    """Vérifie que Qu × PU ≈ Total (tolérance absolue ET relative)."""
    if qu_val and pu_val and total_val:
        try:
            expected = float(qu_val) * float(pu_val)
            actual   = float(total_val)
            if actual != 0:
                abs_diff = abs(expected - actual)
                rel_diff = abs_diff / abs(actual)
                if abs_diff > TOTAL_TOLERANCE_ABS and rel_diff > TOTAL_TOLERANCE_REL:
                    log.warning(
                        "Total incohérent ligne %d code=%s : %.2f × %.2f = %.2f ≠ %.2f",
                        row_idx, code, qu_val, pu_val, expected, actual,
                    )
                    return {
                        "type":    "error",
                        "color":   "red",
                        "row":     row_idx,
                        "code":    code,
                        "message": (
                            f"Total incohérent : {qu_val} × {pu_val} = "
                            f"{expected:.2f} ≠ {actual:.2f} "
                            f"(écart {abs_diff:.2f} €)"
                        ),
                    }
        except (ValueError, TypeError):
            pass
    return None


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_dpgf(filepath: str) -> tuple[pd.DataFrame, list[dict]]:
    """
    Lit et normalise un fichier DPGF entreprise (XLSX, XLS, XLSB).

    Returns:
        dpgf_df (DataFrame) : DataFrame normalisé
        alerts  (list)      : liste d'alertes
    """
    log.info("Lecture DPGF : %s", filepath)

    ext = os.path.splitext(filepath)[1].lower()
    engine = None
    if ext == ".xls": engine = "xlrd"
    elif ext == ".xlsb": engine = "pyxlsb"
    elif ext in [".xlsx", ".xlsm"]: engine = "openpyxl"

    try:
        # Lecture sans header pour trouver la table
        df_raw = pd.read_excel(filepath, engine=engine, header=None)
        header_row_idx = find_header_row(df_raw)
        
        # Re-lecture avec le bon header
        df_data = pd.read_excel(filepath, engine=engine, skiprows=header_row_idx)
    except Exception as e:
        log.error("Erreur de structure DPGF: %s", e)
        return pd.DataFrame(), [{"type": "error", "color": "red", "row": 0, "code": "", "message": str(e)}]

    alerts     = []
    rows       = []
    current_section_code = ""

    # Les colonnes sont identifiées par position
    # 0: Code, 1: Designation, 2: Quantité, 3: Unité, 4: Px U, 5: Px Tot, 12: Entete (col M)
    for idx_in_df, xl_row in df_data.iterrows():
        row_idx = idx_in_df + header_row_idx + 2  # conversion en 1-indexed Excel row
        
        if len(xl_row) < 6:
            continue

        code_raw   = xl_row.iloc[0]
        desig_raw  = xl_row.iloc[1]
        cc_raw     = xl_row.iloc[2]
        u          = xl_row.iloc[3]
        px_u_raw   = xl_row.iloc[4]
        px_tot_raw = xl_row.iloc[5]
        entete     = xl_row.iloc[12] if len(xl_row) > 12 else None

        code_str  = str(code_raw).strip()  if pd.notna(code_raw)  else ""
        desig_str = str(desig_raw).strip() if pd.notna(desig_raw) else ""
        ent_str   = str(entete).strip()    if pd.notna(entete)    else ""

        has_price = bool(
            (pd.notna(cc_raw) and (isinstance(cc_raw, (int, float)) or str(cc_raw).strip())) and 
            (pd.notna(px_u_raw) and (isinstance(px_u_raw, (int, float)) or str(px_u_raw).strip()))
        )

        row_type = classify_row(code_str, desig_str, ent_str, has_price=has_price)

        if row_type == "section_header":
            current_section_code = code_str

        parent_code = current_section_code if row_type == "recap" else ""

        if row_type in ("article", "sub_section"):
            qu_val,  qu_comment  = _clean_numeric(cc_raw)
            pu_val,  pu_comment  = _clean_numeric(px_u_raw)
            tot_val, tot_comment = _clean_numeric(px_tot_raw)
        else:
            qu_val  = cc_raw     if isinstance(cc_raw,     (int, float)) else 0.0
            pu_val  = px_u_raw   if isinstance(px_u_raw,   (int, float)) else 0.0
            tot_val = px_tot_raw if isinstance(px_tot_raw, (int, float)) else 0.0
            qu_comment = pu_comment = tot_comment = ""

        comments    = [c for c in [qu_comment, pu_comment, tot_comment] if c]
        commentaire = "; ".join(comments) if comments else ""

        if row_type == "article" and code_str:
            if qu_comment or pu_comment or tot_comment:
                kw_found = any(
                    c.lower() in KEYWORDS or c.lower() in KEYWORDS.values()
                    for c in [qu_comment, pu_comment, tot_comment] if c
                )
                alert_type = ("info", "blue") if kw_found else ("warning", "yellow")
                msg = (
                    f"Mot-clé détecté : {commentaire}" if kw_found
                    else f"Texte dans champ numérique : {commentaire}"
                )
                alerts.append({
                    "type":    alert_type[0],
                    "color":   alert_type[1],
                    "row":     row_idx,
                    "code":    code_str,
                    "message": msg,
                })
                log.debug("Alerte %s code=%s : %s", alert_type[0], code_str, msg)

            alert = _check_total_coherence(qu_val, pu_val, tot_val, row_idx, code_str)
            if alert:
                alerts.append(alert)

        rows.append({
            "Code":         code_str,
            "Désignation":  desig_str,
            "Qu.":          qu_val,
            "U":            str(u).strip() if pd.notna(u) else "",
            "Px_U_HT":      pu_val,
            "Px_Tot_HT":    tot_val,
            "Commentaire":  commentaire,
            "Entete":       ent_str,
            "row_type":     row_type,
            "original_row": row_idx,
            "parent_code":  parent_code,
        })

    dpgf_df = pd.DataFrame(rows)
    log.info(
        "DPGF parsé : %d lignes, %d alertes",
        len(dpgf_df), len(alerts),
    )
    return dpgf_df, alerts

