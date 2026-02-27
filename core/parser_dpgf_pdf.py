"""
parser_dpgf_pdf.py — Import de DPGF entreprise au format PDF.

Stratégie à deux niveaux :
  1. pdfplumber  : extraction directe des tableaux (PDF vectoriel avec bordures)
  2. PyMuPDF     : reconstruction par coordonnées X/Y (fallback sans tableau structuré)

Retourne (DataFrame, list[dict]) — interface identique à parse_dpgf().
"""

from __future__ import annotations

import re
from decimal import Decimal

import pandas as pd

# Import des helpers numériques depuis parser_dpgf (source unique de vérité)
from core.parser_dpgf import KEYWORDS, _check_total_coherence, _clean_numeric, _looks_numeric
from core.utils import classify_row
from logger import get_logger

log = get_logger(__name__)

# Magic bytes PDF
PDF_MAGIC = b"%PDF"

# ---------------------------------------------------------------------------
# Mots-clés pour détecter la ligne d'en-tête et mapper les colonnes
# ---------------------------------------------------------------------------

_CODE_EXACT = frozenset({"code", "n°", "n°.", "num", "indice", "ref", "no", "n° de prix"})
_DESIG_SUB = ("signation", "libellé", "libelle")  # sous-chaînes suffisantes
_TOT_SUB = ("px tot", "total ht", "prix tot", "montant ht", "px tot.", "prix tot.")
_PU_SUB = ("px u", "p.u", "prix u", "unitaire", "prix unit")
_QU_SUB = ("quantit", "qté", "qte")
_QU_EXACT = frozenset({"qu.", "q.", "qt.", "qt", "qu"})
_U_EXACT = frozenset({"u", "u.", "unité", "unite"})


# ---------------------------------------------------------------------------
# Détection d'en-tête et mapping de colonnes
# ---------------------------------------------------------------------------


def _find_header_idx(rows: list[list]) -> int | None:
    """Trouve l'index de la ligne d'en-tête dans les 40 premières lignes."""
    for i, row in enumerate(rows[:40]):
        if not row:
            continue
        cells = [str(c or "").strip().lower() for c in row]

        has_code = any(c in _CODE_EXACT or c.startswith("n°") for c in cells[:6])
        has_desig = any(any(s in c for s in _DESIG_SUB) for c in cells[:8])
        has_price = any(any(s in c for s in (*_PU_SUB, *_TOT_SUB)) for c in cells)
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

    _try("tot", lambda c: any(s in c for s in _TOT_SUB))
    _try("pu", lambda c: any(s in c for s in _PU_SUB) or c.strip(". ") == "pu")
    _try("qu", lambda c: any(s in c for s in _QU_SUB) or c.strip(". ") in _QU_EXACT)
    _try("u", lambda c: c.strip(". ") in _U_EXACT)
    _try("desig", lambda c: any(s in c for s in _DESIG_SUB))
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
            "vertical_strategy": strategy,
            "horizontal_strategy": strategy,
        }
        with pdfplumber.open(filepath) as pdf:
            for page in pdf.pages:
                for tbl in page.extract_tables(settings):
                    if tbl:
                        rows.extend(tbl)
        return rows

    def _extract_explicit() -> list[list]:
        rows: list[list] = []
        with pdfplumber.open(filepath) as pdf:
            if not pdf.pages:
                return []
            tables = pdf.pages[0].find_tables(
                {"vertical_strategy": "lines", "horizontal_strategy": "lines"}
            )
            if not tables:
                return []
            # Déduire les délimitations des colonnes à partir du 1er tableau trouvé
            first_table = tables[0]
            v_edges = sorted(
                list(set([v[0] for v in first_table.cells] + [v[2] for v in first_table.cells]))
            )
            settings = {
                "vertical_strategy": "explicit",
                "explicit_vertical_lines": v_edges,
                "horizontal_strategy": "text",
            }
            for page in pdf.pages:
                for tbl in page.extract_tables(settings):
                    if tbl:
                        rows.extend(tbl)
        return rows

    try:
        # Essai lattice (bordures)
        rows_lines = _extract_with("lines")
        # Essai explicit vertical (utile quand les bordures s'arrêtent au milieu d'une page, ex ECMB SAS)
        rows_explicit = _extract_explicit()

        # On choisit la méthode extrayant le plus de données structurées
        best_rows = rows_lines
        if rows_explicit and len(rows_explicit) > len(rows_lines) * 1.5:
            if _find_header_idx(rows_explicit) is not None:
                log.info(
                    "pdfplumber explicit : %d lignes (vs %d lines)",
                    len(rows_explicit),
                    len(rows_lines),
                )
                best_rows = rows_explicit

        if best_rows and _find_header_idx(best_rows) is not None:
            if best_rows is rows_lines:
                log.info("pdfplumber lattice : %d lignes", len(best_rows))
            return best_rows

        # Essai stream (alignement texte)
        rows = _extract_with("text")
        if rows and _find_header_idx(rows) is not None:
            log.info("pdfplumber stream : %d lignes", len(rows))
            return rows

        return best_rows or rows or None

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
            left = (col_x[i - 1] + col_x[i]) / 2 if i > 0 else 0.0
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
# Pré-traitement : éclatement des cellules multi-lignes pdfplumber
# ---------------------------------------------------------------------------


def _explode_multiline_rows(rows: list[list]) -> list[list]:
    """Éclate les lignes dont les cellules contiennent des \\n en lignes individuelles.

    Certains PDFs ont leurs données compactées : pdfplumber retourne une seule
    ligne pdfplumber dont chaque cellule contient N valeurs séparées par \\n.
    Cette fonction détecte ce cas et produit N lignes individuelles.

    Alignement intelligent : Code et Désignation ont N sous-valeurs et s'alignent
    1:1. Les colonnes numériques (Qu/PU/Tot) avec M < N sous-valeurs sont
    distribuées aux M lignes dont le code est le plus profond (le plus de
    dot-segments), ce qui correspond aux articles/feuilles de l'arborescence.
    """
    # Déterminer le mapping de colonnes depuis l'en-tête
    header_idx = _find_header_idx(rows)
    code_col_idx: int | None = None
    if header_idx is not None:
        col_map = _map_cols(rows[header_idx])
        code_col_idx = col_map.get("code")

    result: list[list] = []
    for row in rows:
        if not row:
            result.append(row)
            continue

        # Compter combien de cellules contiennent un \n
        has_nl = sum(1 for c in row if isinstance(c, str) and "\n" in c)
        if has_nl < 2:
            result.append(row)
            continue

        # Éclater chaque cellule par \n
        split_cells: list[list[str | None]] = []
        max_parts = 1
        for c in row:
            if isinstance(c, str) and "\n" in c:
                parts = c.split("\n")
                split_cells.append(parts)
                max_parts = max(max_parts, len(parts))
            else:
                split_cells.append([c])

        # N = nombre de sous-lignes de référence (Code / Desig — toujours le plus grand)
        n_ref = max_parts

        # Calculer la profondeur (nb segments) de chaque code
        code_depths: list[int] = []
        if code_col_idx is not None and code_col_idx < len(split_cells):
            code_parts = split_cells[code_col_idx]
            for cval in code_parts:
                c_clean = re.sub(r"\s+", "", str(cval or "")).replace(",", ".")
                segs = [s for s in c_clean.split(".") if s.strip()]
                code_depths.append(len(segs))
            # Padder si code a moins de sous-valeurs que max_parts
            while len(code_depths) < max_parts:
                code_depths.append(0)
        else:
            code_depths = list(range(max_parts))  # fallback : identité

        # Cache d'alignement : pour un nombre donné M de sous-valeurs,
        # quels indices parmi les N lignes reçoivent ces M valeurs ?
        _alignment_cache: dict[int, list[int]] = {}

        def _get_target_indices(
            m: int,
            _alignment_cache=_alignment_cache,
            n_ref=n_ref,
            code_depths=code_depths,
        ) -> list[int]:
            """Retourne les m indices de lignes qui doivent recevoir les m sous-valeurs."""
            if m in _alignment_cache:
                return _alignment_cache[m]
            if m >= n_ref:
                # Autant ou plus de valeurs que de lignes → 1:1
                indices = list(range(n_ref))
            else:
                # Sélectionner les m codes les plus profonds (feuilles de l'arbre)
                # En cas d'égalité de profondeur, garder l'ordre d'apparition
                indexed = [(depth, idx) for idx, depth in enumerate(code_depths[:n_ref])]
                # Tri stable par profondeur DESC — les plus profonds en premier
                indexed.sort(key=lambda x: -x[0])
                top_m = indexed[:m]
                # Re-trier par position originale pour garder l'ordre séquentiel
                indices = sorted(idx for _, idx in top_m)
            _alignment_cache[m] = indices
            return indices

        # Générer n_ref lignes individuelles
        for i in range(n_ref):
            new_row: list[str | None] = []
            for parts in split_cells:
                m = len(parts)
                if m >= n_ref:
                    # Colonne complète → alignement 1:1
                    new_row.append(parts[i] if i < m else None)
                else:
                    # Colonne courte → distribuer aux indices cibles
                    targets = _get_target_indices(m)
                    if i in targets:
                        pos = targets.index(i)
                        new_row.append(parts[pos] if pos < m else None)
                    else:
                        new_row.append(None)
            result.append(new_row)

    return result


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
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "row": 0,
                "code": "",
                "message": (
                    "En-tête non trouvé dans le PDF "
                    "(colonnes Code / Désignation / Prix introuvables). "
                    "Vérifiez que le document est bien un DPGF."
                ),
            }
        )
        return pd.DataFrame()

    col_map = _map_cols(rows[header_idx])
    if "desig" not in col_map and "code" not in col_map:
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "row": 0,
                "code": "",
                "message": "Colonnes Code / Désignation non identifiées dans le PDF.",
            }
        )
        return pd.DataFrame()

    result_rows: list[dict] = []
    current_section_code = ""

    for offset, row in enumerate(rows[header_idx + 1 :]):
        if not row:
            continue

        def _get(key: str, row=row) -> object:
            idx = col_map.get(key)
            if idx is None or idx >= len(row):
                return None
            v = row[idx]
            return v if v is not None else None

        code_raw = _get("code")
        desig_raw = _get("desig")
        qu_raw = _get("qu")
        u_raw = _get("u")
        pu_raw = _get("pu")
        tot_raw = _get("tot")

        code_str = str(code_raw or "").strip()
        # Supprimer les espaces internes dans les codes — artefact courant de
        # l'extraction PDF (PyMuPDF) où "02.1.1" devient "02 .1.1" car les
        # segments sont des mots séparés dans le flux PDF.
        # Un code DPGF légitime ne contient jamais d'espace.
        code_str = re.sub(r"\s+", "", code_str)
        desig_str = str(desig_raw or "").strip()

        if not code_str and not desig_str:
            continue

        has_price = _looks_numeric(qu_raw) and _looks_numeric(pu_raw)
        row_type = classify_row(code_str, desig_str, "", has_price=has_price)

        if row_type == "section_header":
            current_section_code = code_str
        parent_code = current_section_code if row_type == "recap" else ""

        if row_type in ("article", "sub_section"):
            qu_val, qu_cmt = _clean_numeric(qu_raw)
            pu_val, pu_cmt = _clean_numeric(pu_raw)
            tot_val, tot_cmt = _clean_numeric(tot_raw)
        else:
            qu_val = _safe_decimal(qu_raw)
            pu_val = _safe_decimal(pu_raw)
            tot_val = _safe_decimal(tot_raw)
            qu_cmt = pu_cmt = tot_cmt = ""

        # Filtrer les symboles monétaires résiduels des commentaires (artefact PDF)
        _curr = {"€", "$", "£", "eur", "usd"}
        comments = [c for c in [qu_cmt, pu_cmt, tot_cmt] if c and c.strip().lower() not in _curr]
        commentaire = "; ".join(comments) if comments else ""
        u_str = str(u_raw or "").strip()

        # Alerte mots-clés dans champs numériques (même logique que parser_dpgf)
        # Les symboles monétaires résiduels sont déjà filtrés par _curr ci-dessus.
        qu_cmt_clean = qu_cmt if qu_cmt.strip().lower() not in _curr else ""
        pu_cmt_clean = pu_cmt if pu_cmt.strip().lower() not in _curr else ""
        tot_cmt_clean = tot_cmt if tot_cmt.strip().lower() not in _curr else ""
        if row_type == "article" and code_str:
            alert = _check_total_coherence(
                qu_val, pu_val, tot_val, header_idx + 1 + offset + 1, code_str
            )
            if alert:
                alerts.append(alert)
                if commentaire:
                    commentaire += f" ; {alert['short_error']}"
                else:
                    commentaire = f"⚠️ {alert['short_error']}"

            if qu_cmt_clean or pu_cmt_clean or tot_cmt_clean:
                kw_found = any(
                    c.lower() in KEYWORDS or c.lower() in KEYWORDS.values()
                    for c in [qu_cmt, pu_cmt, tot_cmt]
                    if c
                )
                atype = ("info", "blue") if kw_found else ("warning", "yellow")
                msg = (
                    f"Mot-clé détecté : {commentaire}"
                    if kw_found
                    else f"Texte dans champ numérique : {commentaire}"
                )
                alerts.append(
                    {
                        "type": atype[0],
                        "color": atype[1],
                        "row": header_idx + 1 + offset + 1,
                        "code": code_str,
                        "message": msg,
                    }
                )

        result_rows.append(
            {
                "Code": code_str,
                "Désignation": desig_str,
                "Qu.": qu_val,
                "U": u_str,
                "Px_U_HT": pu_val,
                "Px_Tot_HT": tot_val,
                "Commentaire": commentaire,
                "Entete": "",  # Colonne absente dans les PDF
                "row_type": row_type,
                "original_row": header_idx + 1 + offset + 1,
                "parent_code": parent_code,
            }
        )

    if not result_rows:
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "row": 0,
                "code": "",
                "message": "Aucune ligne de données extraite du PDF.",
            }
        )

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
    rows = _extract_pdfplumber(filepath)
    source = "pdfplumber"

    # Niveau 2 : PyMuPDF si pdfplumber n'a pas trouvé d'en-tête valide
    if not rows or _find_header_idx(rows) is None:
        log.info("pdfplumber insuffisant — fallback PyMuPDF")
        rows = _extract_pymupdf(filepath)
        source = "PyMuPDF"
    else:
        # pdfplumber : exploser les cellules multi-lignes avant normalisation
        rows = _explode_multiline_rows(rows)
        log.info("Post-explosion pdfplumber : %d lignes", len(rows))

    if not rows:
        return pd.DataFrame(), [
            {
                "type": "error",
                "color": "red",
                "row": 0,
                "code": "",
                "message": (
                    "Impossible d'extraire les données du PDF. "
                    "Vérifiez que le fichier est un PDF textuel (non scanné) "
                    "et qu'il contient bien un tableau DPGF."
                ),
            }
        ]

    log.info("Source %s : %d lignes brutes extraites", source, len(rows))

    if source == "PyMuPDF":
        alerts.append(
            {
                "type": "info",
                "color": "blue",
                "row": 0,
                "code": "",
                "message": (
                    "PDF sans tableau structuré — extraction par coordonnées X/Y (PyMuPDF). "
                    "Vérifiez que les données extraites sont correctes."
                ),
            }
        )

    df = _normalize_rows(rows, alerts)
    log.info("DPGF PDF parsé : %d lignes, %d alertes", len(df), len(alerts))
    return df, alerts
