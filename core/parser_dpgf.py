"""
parser_dpgf.py — Normalisation et extraction des annotations des DPGF entreprise.

Gère les textes dans les cellules numériques (SANS OBJET, COMPRIS, nc, P-M),
extrait les annotations en colonne Commentaire, et détecte les erreurs.
"""

import pandas as pd
import openpyxl
import re


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

KEYWORDS = {
    "sans objet": "so",
    "compris": "compris",
    "nc": "nc",
    "p-m": "pm",
    "inclus": "inclus",
    "néant": "néant",
    "so": "so",
    "pm": "pm",
}

TOTAL_TOLERANCE = 0.02  # tolérance pour Qu × PU ≈ Total


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _find_header_row(ws):
    """
    Trouve la ligne d'en-tête dans un DPGF entreprise.
    Cherche 'Code' + 'Désignation' + ('Cc' ou 'Qu.').
    """
    for row_idx in range(1, min(ws.max_row + 1, 20)):
        a_val = ws.cell(row=row_idx, column=1).value
        b_val = ws.cell(row=row_idx, column=2).value
        if (
            a_val and str(a_val).strip().lower() == "code"
            and b_val and "signation" in str(b_val).strip().lower()
        ):
            return row_idx
    raise ValueError("Impossible de trouver la ligne d'en-tête du DPGF.")


def _clean_numeric(value):
    """
    Nettoie une valeur potentiellement numérique.
    Retourne (nombre_float, texte_annotion) ou (None, texte_brut).
    """
    if value is None:
        return 0.0, ""

    if isinstance(value, (int, float)):
        return float(value), ""

    text = str(value).strip()
    if not text:
        return 0.0, ""

    # Vérifier si c'est un mot-clé connu
    text_lower = text.lower().strip()
    for keyword, abbrev in KEYWORDS.items():
        if keyword in text_lower:
            return 0.0, abbrev

    # Essayer d'extraire un nombre
    # Format français : remplacer virgule par point, supprimer espaces
    cleaned = text.replace(" ", "").replace("\u00a0", "")
    cleaned = cleaned.replace(",", ".")

    # Chercher un nombre dans le texte
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    if match:
        number = float(match.group())
        # Le reste du texte est une annotation
        remaining = re.sub(r"-?\d+(?:\.\d+)?", "", cleaned).strip()
        remaining = remaining.strip("()[]{}/ ")
        return number, remaining

    # Texte pur sans nombre
    return 0.0, text


def _is_total_row(code_val):
    """Vérifie si la ligne est une ligne de total."""
    if not code_val:
        return False
    return str(code_val).strip().lower().startswith("total")


# ---------------------------------------------------------------------------
# Alertes
# ---------------------------------------------------------------------------

def _check_total_coherence(qu_val, pu_val, total_val, row_idx, code):
    """Vérifie que Qu × PU ≈ Total. Retourne une alerte si incohérent."""
    if qu_val and pu_val and total_val:
        try:
            expected = float(qu_val) * float(pu_val)
            actual = float(total_val)
            if actual != 0 and abs(expected - actual) > TOTAL_TOLERANCE:
                return {
                    "type": "error",
                    "color": "red",
                    "row": row_idx,
                    "code": code,
                    "message": (
                        f"Total incohérent : {qu_val} × {pu_val} = "
                        f"{expected:.2f} ≠ {actual:.2f}"
                    ),
                }
        except (ValueError, TypeError):
            pass
    return None


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------

def parse_dpgf(filepath):
    """
    Lit et normalise un fichier DPGF entreprise (.xlsx).

    Retourne :
        dpgf_df : DataFrame avec colonnes
            [Code, Désignation, Qu., U, Px_U_HT, Px_Tot_HT, Commentaire,
             Entete, row_type, original_row]
        alerts  : liste de dicts décrivant les anomalies détectées
    """
    wb = openpyxl.load_workbook(filepath, data_only=True)
    ws = wb.active

    header_row = _find_header_row(ws)
    alerts = []

    rows = []
    for row_idx in range(header_row + 1, ws.max_row + 1):
        code = ws.cell(row=row_idx, column=1).value
        designation = ws.cell(row=row_idx, column=2).value
        cc_raw = ws.cell(row=row_idx, column=3).value      # Cc (quantité)
        u = ws.cell(row=row_idx, column=4).value
        px_u_raw = ws.cell(row=row_idx, column=5).value     # Px U
        px_tot_raw = ws.cell(row=row_idx, column=6).value   # Px Total
        entete = ws.cell(row=row_idx, column=13).value      # col M

        # Déterminer le type de ligne
        if _is_total_row(code):
            row_type = "total"
        elif entete and "Recap" in str(entete):
            row_type = "recap"
        elif code and designation:
            row_type = "data"
        elif not code and not designation:
            row_type = "empty"
        else:
            row_type = "other"

        # Normaliser les valeurs numériques
        qu_val, qu_comment = _clean_numeric(cc_raw)
        pu_val, pu_comment = _clean_numeric(px_u_raw)
        tot_val, tot_comment = _clean_numeric(px_tot_raw)

        # Construire le commentaire consolidé
        comments = [c for c in [qu_comment, pu_comment, tot_comment] if c]
        commentaire = "; ".join(comments) if comments else ""

        # Générer les alertes pour les lignes de données
        code_str = str(code).strip() if code else ""

        if row_type == "data" and code_str:
            # Alerte texte dans valeur numérique
            if qu_comment or pu_comment or tot_comment:
                kw_found = any(
                    c.lower() in KEYWORDS or c.lower() in KEYWORDS.values()
                    for c in [qu_comment, pu_comment, tot_comment]
                    if c
                )
                if kw_found:
                    alerts.append({
                        "type": "info",
                        "color": "blue",
                        "row": row_idx,
                        "code": code_str,
                        "message": f"Mot-clé détecté : {commentaire}",
                    })
                else:
                    alerts.append({
                        "type": "warning",
                        "color": "yellow",
                        "row": row_idx,
                        "code": code_str,
                        "message": f"Texte dans champ numérique : {commentaire}",
                    })

            # Alerte total incohérent
            alert = _check_total_coherence(qu_val, pu_val, tot_val, row_idx, code_str)
            if alert:
                alerts.append(alert)

        rows.append({
            "Code": code_str,
            "Désignation": str(designation).strip() if designation else "",
            "Qu.": qu_val,
            "U": str(u).strip() if u else "",
            "Px_U_HT": pu_val,
            "Px_Tot_HT": tot_val,
            "Commentaire": commentaire,
            "Entete": str(entete).strip() if entete else "",
            "row_type": row_type,
            "original_row": row_idx,
        })

    dpgf_df = pd.DataFrame(rows)
    wb.close()
    return dpgf_df, alerts
