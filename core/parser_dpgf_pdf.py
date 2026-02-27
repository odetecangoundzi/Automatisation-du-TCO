"""
parser_dpgf_pdf.py — Import de DPGF entreprise au format PDF.

Stratégie à deux niveaux :
  1. pdfplumber  : extraction directe des tableaux (PDF vectoriel avec bordures)
  2. PyMuPDF     : reconstruction par coordonnées X/Y (fallback sans tableau structuré)

Retourne (DataFrame, list[dict]) — interface identique à parse_dpgf().
"""

from __future__ import annotations

from decimal import Decimal

import pandas as pd

# Import des helpers numériques depuis parser_dpgf (source unique de vérité)
from core.parser_dpgf import KEYWORDS, _clean_numeric, _looks_numeric
from core.utils import classify_row
from logger import get_logger

log = get_logger(__name__)

# Magic bytes PDF
PDF_MAGIC = b"%PDF"

# ---------------------------------------------------------------------------
# Mots-clés pour détecter la ligne d'en-tête et mapper les colonnes
# ---------------------------------------------------------------------------

_CODE_EXACT = frozenset({"code", "n°", "n°.", "num", "indice", "ref", "no", "n° de prix"})
_DESIG_SUB  = ("signation", "libellé", "libelle")   # sous-chaînes suffisantes
_TOT_SUB    = ("px tot", "total ht", "prix tot", "montant ht", "px tot.", "prix tot.")
_PU_SUB     = ("px u", "p.u", "prix u", "unitaire", "prix unit")
_QU_SUB     = ("quantit", "qté", "qte")
_QU_EXACT   = frozenset({"qu.", "q.", "qt.", "qt", "qu"})
_U_EXACT    = frozenset({"u", "u.", "unité", "unite"})


# ---------------------------------------------------------------------------
# Détection d'en-tête et mapping de colonnes
# ---------------------------------------------------------------------------


def _find_header_idx(rows: list[list]) -> int | None:
    """Trouve l'index de la ligne d'en-tête dans les 40 premières lignes."""
    for i, row in enumerate(rows[:40]):
        if not row:
            continue
        cells = [str(c or "").strip().lower() for c in row]

        has_code  = any(c in _CODE_EXACT or c.startswith("n°") for c in cells[:6])
        has_desig = any(any(s in c for s in _DESIG_SUB) for c in cells[:8])
        has_price = any(
            any(s in c for s in (*_PU_SUB, *_TOT_SUB))
            for c in cells
        )
        if (has_code or has_desig) and has_price:
            return i
    return None


def _map_cols(header_row: list) -> dict[str, int]:
    """
    Mappe colonnes DPGF → indices, gauche → droite.
    Priorité : tot > pu > qu > u > desig > code (du plus spécifique au moins).
    """
    cells = [str(c or "").strip().lower() for c in header_row]
    result: dict[str, int] = {}
    taken: set[int] = set()

    def _try(key: str, matcher) -> None:
        if key in result:
            return
        for i, c in enumerate(cells):
            if i not in taken and matcher(c):
                result[key] = i
                taken.add(i)
                break

    _try("tot",  lambda c: any(s in c for s in _TOT_SUB))
    _try("pu",   lambda c: any(s in c for s in _PU_SUB) or c.strip(". ") == "pu")
    _try("qu",   lambda c: any(s in c for s in _QU_SUB) or c.strip(". ") in _QU_EXACT)
    _try("u",    lambda c: c.strip(". ") in _U_EXACT)
    _try("desig",lambda c: any(s in c for s in _DESIG_SUB))
    _try("code", lambda c: c.strip(". ") in _CODE_EXACT or c.startswith("n°"))

    return result


# ---------------------------------------------------------------------------
# Niveau 1 : pdfplumber
# ---------------------------------------------------------------------------


def _extract_pdfplumber(filepath: str) -> list[list] | None:
    """
    Extrait les tableaux du PDF via pdfplumber.

    Essaie deux stratégies :
      1. lattice  : tableaux avec bordures visibles
      2. stream   : tableaux par alignement de texte
    """
    try:
        import pdfplumber  # noqa: PLC0415
    except ImportError:
        log.warning("pdfplumber non installé — fallback PyMuPDF")
        return None

    def _extract_with(strategy: str) -> list[list]:
        rows: list[list] = []
        settings = {
            "vertical_strategy":   strategy,
            "horizontal_strategy": strategy,
        }
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                for tbl in page.extract_tables(settings):
                    if tbl:
                        rows.extend(tbl)
        return rows

    try:
        # Essai lattice (bordures)
        rows = _extract_with("lines")
        if rows and _find_header_idx(rows) is not None:
            log.info("pdfplumber lattice : %d lignes", len(rows))
            return rows

        # Essai stream (alignement texte)
        rows = _extract_with("text")
        if rows and _find_header_idx(rows) is not None:
            log.info("pdfplumber stream : %d lignes", len(rows))
            return rows

        return rows or None

    except Exception as exc:
        log.warning("pdfplumber erreur : %s", exc)
        return None


# ---------------------------------------------------------------------------
# Niveau 2 : PyMuPDF (fallback)
# ---------------------------------------------------------------------------


def _extract_pymupdf(filepath: str) -> list[list] | None:
    """
    Reconstruit un tableau depuis les coordonnées X/Y des mots (PyMuPDF).

    Algorithme :
      1. Extraire les mots avec (x0, y0, word)
      2. Grouper par ligne (y0 similaire ± Y_TOL)
      3. Trouver la ligne d'en-tête → déduire les bornes de colonnes
      4. Assigner chaque mot à la colonne la plus proche (mi-chemin entre colonnes)
    """
    try:
        import fitz  # noqa: PLC0415
    except ImportError:
        log.warning("PyMuPDF (fitz) non installé")
        return None

    try:
        doc = fitz.open(filepath)
        # Récupérer tous les mots de toutes les pages
        page_word_lines: list[list[tuple[float, str]]] = []

        for page in doc:
            words = page.get_text("words")  # (x0,y0,x1,y1,word,block,line,word_no)
            if not words:
                continue

            Y_TOL = 4.0
            lines_y: dict[float, list[tuple[float, str]]] = {}
            for x0, y0, _x1, _y1, word, *_ in words:
                matched = next((ly for ly in lines_y if abs(ly - y0) <= Y_TOL), None)
                if matched is None:
                    lines_y[y0] = []
                    matched = y0
                lines_y[matched].append((x0, word))

            for y in sorted(lines_y):
                page_word_lines.append(sorted(lines_y[y], key=lambda w: w[0]))

        doc.close()

        if not page_word_lines:
            return None

        # Trouver la ligne d'en-tête
        raw_rows = [[w for _, w in line] for line in page_word_lines]
        header_idx = _find_header_idx(raw_rows)
        if header_idx is None:
            return None

        # Positions X des colonnes depuis l'en-tête
        col_x = [x for x, _ in page_word_lines[header_idx]]
        n_cols = len(col_x)
        if n_cols == 0:
            return None

        # Bornes de colonnes : mi-chemin entre deux x consécutifs
        bounds: list[tuple[float, float]] = []
        for i in range(n_cols):
            left  = (col_x[i - 1] + col_x[i]) / 2 if i > 0 else 0.0
            right = (col_x[i] + col_x[i + 1]) / 2 if i < n_cols - 1 else float("inf")
            bounds.append((left, right))

        def _col_for(x0: float) -> int:
            for ci, (left_bound, right_bound) in enumerate(bounds):
                if left_bound <= x0 < right_bound:
                    return ci
            # Hors bornes → colonne la plus proche
            return min(range(n_cols), key=lambda i: abs(col_x[i] - x0))

        # Reconstruction des lignes
        result: list[list[str]] = []
        for line in page_word_lines:
            row = [""] * n_cols
            for x0, word in line:
                ci = _col_for(x0)
                row[ci] = (row[ci] + " " + word).strip()
            result.append(row)

        log.info("PyMuPDF : %d lignes reconstruites (%d colonnes)", len(result), n_cols)
        return result

    except Exception as exc:
        log.warning("PyMuPDF erreur : %s", exc)
        return None


# ---------------------------------------------------------------------------
# Normalisation lignes brutes → DataFrame
# ---------------------------------------------------------------------------


def _safe_decimal(val: object) -> Decimal:
    """Convertit une valeur en Decimal sans lever d'exception."""
    if val is None:
        return Decimal("0.0")
    s = str(val).strip().replace("\u00a0", "").replace(" ", "").replace(",", ".")
    if not s:
        return Decimal("0.0")
    try:
        return Decimal(s)
    except Exception:
        return Decimal("0.0")


def _normalize_rows(rows: list[list], alerts: list[dict]) -> pd.DataFrame:
    """Convertit les lignes brutes extraites en DataFrame normalisé."""
    header_idx = _find_header_idx(rows)
    if header_idx is None:
        alerts.append({
            "type": "error", "color": "red", "row": 0, "code": "",
            "message": (
                "En-tête non trouvé dans le PDF "
                "(colonnes Code / Désignation / Prix introuvables). "
                "Vérifiez que le document est bien un DPGF."
            ),
        })
        return pd.DataFrame()

    col_map = _map_cols(rows[header_idx])
    if "desig" not in col_map and "code" not in col_map:
        alerts.append({
            "type": "error", "color": "red", "row": 0, "code": "",
            "message": "Colonnes Code / Désignation non identifiées dans le PDF.",
        })
        return pd.DataFrame()

    result_rows: list[dict] = []
    current_section_code = ""

    for offset, row in enumerate(rows[header_idx + 1:]):
        if not row:
            continue

        def _get(key: str, row=row) -> object:
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return None
            v = row[idx]
            return v if v is not None else None

        code_raw  = _get("code")
        desig_raw = _get("desig")
        qu_raw    = _get("qu")
        u_raw     = _get("u")
        pu_raw    = _get("pu")
        tot_raw   = _get("tot")

        code_str  = str(code_raw  or "").strip()
        desig_str = str(desig_raw or "").strip()

        if not code_str and not desig_str:
            continue

        has_price = _looks_numeric(qu_raw) and _looks_numeric(pu_raw)
        row_type  = classify_row(code_str, desig_str, "", has_price=has_price)

        if row_type == "section_header":
            current_section_code = code_str
        parent_code = current_section_code if row_type == "recap" else ""

        if row_type in ("article", "sub_section"):
            qu_val,  qu_cmt  = _clean_numeric(qu_raw)
            pu_val,  pu_cmt  = _clean_numeric(pu_raw)
            tot_val, tot_cmt = _clean_numeric(tot_raw)
        else:
            qu_val  = _safe_decimal(qu_raw)
            pu_val  = _safe_decimal(pu_raw)
            tot_val = _safe_decimal(tot_raw)
            qu_cmt = pu_cmt = tot_cmt = ""

        comments    = [c for c in [qu_cmt, pu_cmt, tot_cmt] if c]
        commentaire = "; ".join(comments) if comments else ""
        u_str       = str(u_raw or "").strip()

        # Alerte mots-clés dans champs numériques (même logique que parser_dpgf)
        if row_type == "article" and code_str and (qu_cmt or pu_cmt or tot_cmt):
            kw_found   = any(
                c.lower() in KEYWORDS or c.lower() in KEYWORDS.values()
                for c in [qu_cmt, pu_cmt, tot_cmt]
                if c
            )
            atype  = ("info", "blue") if kw_found else ("warning", "yellow")
            msg    = (
                f"Mot-clé détecté : {commentaire}"
                if kw_found
                else f"Texte dans champ numérique : {commentaire}"
            )
            alerts.append({
                "type": atype[0], "color": atype[1],
                "row": header_idx + 1 + offset + 1,
                "code": code_str, "message": msg,
            })

        result_rows.append({
            "Code":        code_str,
            "Désignation": desig_str,
            "Qu.":         qu_val,
            "U":           u_str,
            "Px_U_HT":     pu_val,
            "Px_Tot_HT":   tot_val,
            "Commentaire": commentaire,
            "Entete":      "",   # Colonne absente dans les PDF
            "row_type":    row_type,
            "original_row": header_idx + 1 + offset + 1,
            "parent_code": parent_code,
        })

    if not result_rows:
        alerts.append({
            "type": "warning", "color": "orange", "row": 0, "code": "",
            "message": "Aucune ligne de données extraite du PDF.",
        })

    return pd.DataFrame(result_rows) if result_rows else pd.DataFrame()


# ---------------------------------------------------------------------------
# Point d'entrée principal
# ---------------------------------------------------------------------------


def parse_dpgf_pdf(filepath: str) -> tuple[pd.DataFrame, list[dict]]:
    """
    Lit et normalise un fichier DPGF entreprise au format PDF.

    Stratégie :
      1. pdfplumber — tableaux avec bordures (lattice) puis par texte (stream)
      2. PyMuPDF    — reconstruction X/Y si pdfplumber ne trouve pas d'en-tête

    Returns:
        dpgf_df : DataFrame normalisé (même structure que parse_dpgf)
        alerts  : liste d'alertes
    """
    log.info("Lecture DPGF PDF : %s", filepath)
    alerts: list[dict] = []

    # Niveau 1 : pdfplumber
    rows  = _extract_pdfplumber(filepath)
    source = "pdfplumber"

    # Niveau 2 : PyMuPDF si pdfplumber n'a pas trouvé d'en-tête valide
    if not rows or _find_header_idx(rows) is None:
        log.info("pdfplumber insuffisant — fallback PyMuPDF")
        rows  = _extract_pymupdf(filepath)
        source = "PyMuPDF"

    if not rows:
        return pd.DataFrame(), [{
            "type": "error", "color": "red", "row": 0, "code": "",
            "message": (
                "Impossible d'extraire les données du PDF. "
                "Vérifiez que le fichier est un PDF textuel (non scanné) "
                "et qu'il contient bien un tableau DPGF."
            ),
        }]

    log.info("Source %s : %d lignes brutes extraites", source, len(rows))

    if source == "PyMuPDF":
        alerts.append({
            "type": "info", "color": "blue", "row": 0, "code": "",
            "message": (
                "PDF sans tableau structuré — extraction par coordonnées X/Y (PyMuPDF). "
                "Vérifiez que les données extraites sont correctes."
            ),
        })

    df = _normalize_rows(rows, alerts)
    log.info("DPGF PDF parsé : %d lignes, %d alertes", len(df), len(alerts))
    return df, alerts
