"""
exporter.py — Export du TCO fusionné en fichier Excel formaté.

Génère un fichier .xlsx avec :
- Headers groupés par entreprise (merged cells)
- Mise en forme (gras, freeze pane, largeur auto)
- Coloration selon alertes
- Lignes section_header et recap mises en évidence
- Support export via BytesIO (pas de sauvegarde disque obligatoire)
"""

from __future__ import annotations

import io
import re
from typing import TYPE_CHECKING

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from logger import get_logger

if TYPE_CHECKING:
    import pandas as pd

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Constantes de style
# ---------------------------------------------------------------------------

FONT_HEADER = Font(name="Tahoma", bold=True, size=10, color="FFFFFF")
FONT_HEADER_COMPANY = Font(name="Tahoma", bold=True, size=10, color="FFFFFF")
FONT_SECTION = Font(name="Tahoma", bold=True, size=11, color="AC2C18")  # rouge foncé — référence
FONT_RECAP = Font(name="Tahoma", bold=True, size=11, color="000000")  # noir gras
FONT_TOTAL = Font(name="Tahoma", bold=True, size=11, color="000000")  # noir gras
FONT_GRAND_TOTAL = Font(name="Tahoma", bold=True, size=11, color="FFFFFF")  # blanc sur fond foncé
FONT_DATA = Font(name="Tahoma", size=9, color="000000")
FONT_SUB_SECTION = Font(name="Tahoma", bold=True, size=9, color="314E85")  # bleu foncé — référence

FILL_HEADER = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
FILL_COMPANY_COLORS = [
    PatternFill(start_color="548235", end_color="548235", fill_type="solid"),
    PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid"),
    PatternFill(start_color="843C0C", end_color="843C0C", fill_type="solid"),
    PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid"),
    PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"),
]

# Lignes de données : fond blanc pur (conforme référence — hiérarchie via couleur police)
# Format ARGB 8 chars : "FFFFFFFF" = blanc opaque — correspond à fgColor.rgb de la référence
FILL_WHITE = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
FILL_SECTION = FILL_WHITE
FILL_RECAP = FILL_WHITE
FILL_RECAP_SUMMARY = PatternFill(  # bleu clair — distingue la section récapitulatif
    start_color="D6E4F7", end_color="D6E4F7", fill_type="solid"
)
FILL_TOTAL_LINE = FILL_WHITE
FILL_RECAP_HEADER = PatternFill(  # marine foncé — bandeau titre récapitulatif
    start_color="17375E", end_color="17375E", fill_type="solid"
)
FILL_SUB_SECTION = FILL_WHITE

# Titres principaux (sub_section sans prix = ex : BATIMENT F)
FONT_MAIN_TITLE = Font(name="Tahoma", bold=True, size=11, color="314E85")  # bleu foncé ref
FILL_MAIN_TITLE = FILL_WHITE

# Totaux généraux — fond sombre + texte blanc (FONT_GRAND_TOTAL)
FILL_MONTANT_HT = PatternFill(
    start_color="1F4E79", end_color="1F4E79", fill_type="solid"
)  # bleu foncé
FILL_TVA = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")  # bleu moyen
FILL_MONTANT_TTC = PatternFill(
    start_color="0D2137", end_color="0D2137", fill_type="solid"
)  # bleu très foncé

FILL_ERROR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_WARNING = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
FILL_NOTE = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
FILL_INFO = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")

FONT_RECAP_HEADER = Font(name="Tahoma", bold=True, size=11, color="FFFFFF")

# Articles alternatifs (code suffixé "A", "AA"…) — orange, conforme modèles de référence
FONT_ALTERNATIVE = Font(name="Tahoma", bold=True, size=9, color="DC9329")

# Récap de section imbriquée (sous-total intermédiaire) — taille 8pt comme la référence
FONT_RECAP_SUB = Font(name="Tahoma", bold=True, size=8, color="000000")

# ---------------------------------------------------------------------------
# Couleurs d'onglet par numéro de lot (palette cyclique — supporte tous les lots)
# ---------------------------------------------------------------------------
_LOT_PALETTE: list[str] = [
    "548235",  # 01 vert
    "843C0C",  # 02 brun
    "1F4E79",  # 03 bleu foncé
    "7030A0",  # 04 violet
    "2E75B6",  # 05 bleu moyen
    "BF8F00",  # 06 or
    "C00000",  # 07 rouge
    "375623",  # 08 vert foncé
    "984807",  # 09 orange foncé
    "2E4057",  # 10 ardoise
    "5C4827",  # 11 brun foncé
    "1C6E52",  # 12 vert émeraude
]

# Correspondances sémantiques historiques (lots 01-06, 10-12)
_SEMANTIC_COLORS: dict[str, str] = {
    "01": "548235",
    "02": "843C0C",
    "03": "1F4E79",
    "04": "7030A0",
    "05": "2E75B6",
    "06": "BF8F00",
    "10": "C00000",
    "11": "375623",
    "12": "984807",
}


def _get_lot_tab_color(lot_num: str) -> str:
    """Retourne la couleur d'onglet pour un numéro de lot quelconque.

    Les lots 01-06 et 10-12 conservent leurs couleurs historiques.
    Les autres lots utilisent la palette cyclique (_LOT_PALETTE).
    """
    if lot_num in _SEMANTIC_COLORS:
        return _SEMANTIC_COLORS[lot_num]
    try:
        return _LOT_PALETTE[(int(lot_num) - 1) % len(_LOT_PALETTE)]
    except (ValueError, TypeError):
        return "2F5496"  # bleu par défaut

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)
# Formats numériques — format exact de la référence
MONEY_FORMAT = r"###,###,###,##0.00\ \€;\-###,###,###,##0.00\ \€;"
QTY_FORMAT = r"###,###,###,##0.00;\-###,###,###,##0.00;"


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _detect_companies(df: pd.DataFrame) -> list[str]:
    """Détecte les noms d'entreprises à partir des colonnes _Qu."""
    companies, seen = [], set()
    for col in df.columns:
        if col.endswith("_Qu."):
            name = col[:-4]
            if name not in seen:
                companies.append(name)
                seen.add(name)
    return companies


def _auto_width(ws, min_width: int = 8, max_width: int = 40) -> None:
    """Ajuste la largeur des colonnes automatiquement."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = min_width
        for cell in col_cells:
            if cell.value:
                # Truncate string representation to check length, avoid huge cells
                val_str = str(cell.value)
                line_len = (
                    max(len(line) for line in val_str.split("\n"))
                    if "\n" in val_str
                    else len(val_str)
                )
                max_len = max(max_len, min(line_len + 2, max_width))
        ws.column_dimensions[col_letter].width = max_len


def _get_alert_fill(color: str) -> PatternFill | None:
    return {"red": FILL_ERROR, "orange": FILL_WARNING, "yellow": FILL_NOTE, "blue": FILL_INFO}.get(
        color
    )


def _get_row_style(row_type: str) -> tuple[Font, PatternFill | None]:
    return {
        "section_header": (FONT_SECTION, FILL_SECTION),
        "recap": (FONT_RECAP, FILL_RECAP),
        "recap_summary": (FONT_RECAP, FILL_RECAP_SUMMARY),
        "total_line": (FONT_TOTAL, FILL_TOTAL_LINE),
        "sub_section": (FONT_SUB_SECTION, FILL_SUB_SECTION),
    }.get(row_type, (FONT_DATA, None))


def fix_freeze_panes(ws, header_rows: int = 2, frozen_cols: int = 0) -> None:
    """
    Garantit que le freeze panes est positionné à la cellule ancre correcte.
    header_rows=2, frozen_cols=0  →  ancre A3
    (lignes 1-2 figées + 0 colonnes figées).
    """
    anchor = f"{get_column_letter(frozen_cols + 1)}{header_rows + 1}"
    ws.freeze_panes = anchor


def fix_merged_cells_crossing_freeze(
    ws,
    header_rows: int = 2,
    frozen_cols: int = 2,
) -> None:
    """
    Supprime toute fusion qui traverse la frontière de freeze panes.
    Les fusions horizontales sont remplacées par centerContinuous pour
    conserver l'effet visuel sans provoquer de chevauchement au scroll.
    """
    to_process = []
    for mr in list(ws.merged_cells.ranges):
        crosses_col = mr.min_col <= frozen_cols < mr.max_col
        crosses_row = mr.min_row <= header_rows < mr.max_row
        if not (crosses_col or crosses_row):
            continue
        pivot = ws.cell(row=mr.min_row, column=mr.min_col)
        to_process.append(
            (
                mr.coord,
                mr.min_row,
                mr.min_col,
                mr.max_row,
                mr.max_col,
                pivot.value,
                pivot.font.copy() if pivot.font else None,
                pivot.fill.copy() if pivot.fill else None,
            )
        )

    for coord, min_row, min_col, max_row, max_col, val, fnt, fll in to_process:
        ws.unmerge_cells(coord)
        for r in range(min_row, max_row + 1):
            for c in range(min_col, max_col + 1):
                cell = ws.cell(row=r, column=c)
                if fnt:
                    cell.font = fnt
                if fll:
                    cell.fill = fll
                if max_row == min_row:  # fusion horizontale → center across
                    cell.alignment = Alignment(horizontal="centerContinuous")
        ws.cell(row=min_row, column=min_col).value = val
        log.debug("Fusion crossing freeze corrigée : %s", coord)


def prevent_text_overflow(
    ws,
    min_row: int = 3,
    max_row: int | None = None,
    min_col: int = 1,
    max_col: int | None = None,
) -> None:
    """
    Garantit qu'aucune cellule du tableau n'est transparente (fill_type=None).
    Un fill blanc sur les cellules vides empêche le texte adjacent de déborder
    horizontalement pendant le scroll (effet "spill" Excel).
    """
    if max_row is None:
        max_row = ws.max_row
    if max_col is None:
        max_col = ws.max_column
    for r in range(min_row, max_row + 1):
        for c in range(min_col, max_col + 1):
            cell = ws.cell(row=r, column=c)
            ft = cell.fill.fill_type if cell.fill else None
            if ft is None or ft == "none":
                cell.fill = FILL_WHITE


def _rows_to_sum_formula(col: str, rows: list[int]) -> str:
    """
    Convertit une liste de numéros de lignes Excel en formule =SUM() avec plages.

    Exemple : [3,4,5,7,8] → '=SUM(F3:F5,F7:F8)'
    Jamais d'énumération de cellules individuelles : toujours des plages contiguës.
    """
    if not rows:
        return "0"
    sorted_rows = sorted(set(rows))
    parts: list[str] = []
    start = end = sorted_rows[0]
    for r in sorted_rows[1:]:
        if r == end + 1:
            end = r
        else:
            parts.append(f"{col}{start}:{col}{end}" if start != end else f"{col}{start}")
            start = end = r
    parts.append(f"{col}{start}:{col}{end}" if start != end else f"{col}{start}")
    return "=SUM(" + ",".join(parts) + ")"


# ---------------------------------------------------------------------------
# Main exporter
# ---------------------------------------------------------------------------


def export_tco(
    merged_df: pd.DataFrame,
    meta: dict,
    output_path: str | None = None,
    alerts: list[dict] | None = None,
    tva_rate: float = 0.20,
) -> str | io.BytesIO:
    """
    Exporte le TCO fusionné en fichier Excel formaté.
    """
    if alerts is None:
        alerts = []

    log.info("Début export Excel. Lignes=%d", len(merged_df))

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = meta.get("sheet_name", "TCO Final")

    # --- Identification du lot : numéro + couleur d'onglet ---
    _lot_raw = ((meta.get("project_info") or {}).get("lot", "") or "").strip()
    _lot_match = re.search(r"\b(\d{2})\b", _lot_raw)
    _lot_num = _lot_match.group(1) if _lot_match else ""
    _tab_color = _get_lot_tab_color(_lot_num) if _lot_num else "2F5496"
    ws.sheet_properties.tabColor = _tab_color
    log.debug("Lot détecté : '%s' → tab color #%s", _lot_num or "?", _tab_color)

    companies = _detect_companies(merged_df)
    log.debug("Entreprises détectées : %s", companies)

    # --- ROW 1-4 : Metadata (MOA, MOE, etc.) ---
    project_info = meta.get("project_info", {})
    # Access metadata potentially passed down from the active project via the caller
    moa = project_info.get("moa", "")
    moe = project_info.get("moe", "")
    devise = project_info.get("devise", "€")

    current_row = 1
    if moa or moe:
        ws.cell(row=current_row, column=1, value="Maître d'Ouvrage :").font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=moa)
        current_row += 1
        ws.cell(row=current_row, column=1, value="Maître d'Œuvre :").font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=moe)
        current_row += 1
        ws.cell(row=current_row, column=1, value="Devise :").font = Font(bold=True)
        ws.cell(row=current_row, column=2, value=devise)
        current_row += 2 # Add a blank line

    header_row_1 = current_row
    header_row_2 = current_row + 1

    # --- Header groupés ---
    ws.cell(row=header_row_1, column=2, value="Etudes")
    ws.cell(row=header_row_1, column=3, value=" Estimation")
    ws.merge_cells(start_row=header_row_1, start_column=3, end_row=header_row_1, end_column=6)
    # Col 1 (A1) incluse : fill obligatoire pour bloquer le débordement de B1 au scroll
    for c in range(1, 7):
        cell = ws.cell(row=header_row_1, column=c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")

    # Badge LOT en première colonne : fond couleur du lot + texte "LOT XX" en blanc
    # Conforme aux modèles de référence qui identifient le lot dès la 1ère cellule.
    if _lot_num:
        _a1 = ws.cell(row=header_row_1, column=1, value=f"LOT {_lot_num}")
        _a1.font = Font(name="Tahoma", bold=True, size=10, color="FFFFFF")
        _a1.fill = PatternFill(start_color=_tab_color, end_color=_tab_color, fill_type="solid")
        _a1.alignment = Alignment(horizontal="center", vertical="center")

    company_start_col = 7
    for comp_idx, comp in enumerate(companies):
        start_col = company_start_col
        end_col = start_col + 3
        ws.cell(row=header_row_1, column=start_col, value=comp)
        ws.merge_cells(start_row=header_row_1, start_column=start_col, end_row=header_row_1, end_column=end_col)
        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=header_row_1, column=c)
            cell.font = FONT_HEADER_COMPANY
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")
        company_start_col = end_col + 1

    # --- Noms de colonnes ---
    base_headers = ["Code", "Désignation", "Qu.", "U", "Px U. HT", "Px Tot HT"]
    for i, header in enumerate(base_headers, 1):
        cell = ws.cell(row=header_row_2, column=i, value=header)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER

    col_offset = 7
    for comp_idx, _comp in enumerate(companies):
        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for i, header in enumerate(["Qu.", "Px U. HT", "Px Tot HT", "Commentaire poste"]):
            cell = ws.cell(row=header_row_2, column=col_offset + i, value=header)
            cell.font = FONT_HEADER
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
        col_offset += 4

    max_col = 6 + len(companies) * 4

    # Index des alertes par code
    alert_by_code = {}
    for alert in alerts:
        code = alert.get("code", "")
        if code:
            alert_by_code.setdefault(code, []).append(alert)

    # --- PRE-COMPUTE : sections actives ---
    # Une section est "active" si elle contient au moins un article/sub_section avec
    # un montant non nul (template OU entreprise) ou une anomalie détectée.
    # Les sections 100% vides (ex : 01.3 … 01.10 sans données) sont ignorées à l'export.
    active_sections: set[str] = set()
    _cur_sec_scan: str | None = None
    for _, _row in merged_df.iterrows():
        _rt = _row["row_type"]
        if _rt == "section_header":
            _cur_sec_scan = str(_row.get("Code", "")).strip()
        elif _rt in ("article", "sub_section") and _cur_sec_scan:
            _code_scan = str(_row.get("Code", "")).strip()
            _has_data = False
            try:
                _has_data = float(_row.get("Px_Tot_HT", 0) or 0) != 0
            except (ValueError, TypeError):
                pass
            if not _has_data:
                for _comp in companies:
                    try:
                        if float(_row.get(f"{_comp}_Px_Tot_HT", 0) or 0) != 0:
                            _has_data = True
                            break
                    except (ValueError, TypeError):
                        pass
            if not _has_data and _code_scan and _code_scan in alert_by_code:
                _has_data = True
            if _has_data:
                active_sections.add(_cur_sec_scan)
    log.debug("Sections actives (%d) : %s", len(active_sections), sorted(active_sections))

    # --- ROWS 3+ : Données ---
    excel_row = 3

    # Tracking pour les formules dynamiques
    section_articles: dict[str, list[int]] = {}  # { '01.1': [row_idx, ...] }
    section_total_row: dict[str, int] = {}  # { '01.1': recap_row_idx }
    section_header_rows: dict[str, int] = {}  # { '01.1': section_header_row_idx }
    recap_summary_rows: list[int] = []  # [row_idx, ...]

    current_section_code: str | None = None
    _skip_section: bool = False  # True si la section courante est vide
    _recap_header_written: bool = False  # True dès que le bandeau RÉCAPITULATIF a été inséré

    # QW-3 : Variables initialisées ici (pas de 'in locals()' fragile)
    ht_row_idx: int | None = None
    tva_row_idx: int | None = None

    for _, row in merged_df.iterrows():
        row_type = row["row_type"]
        if row_type == "empty":
            continue

        code = str(row.get("Code", "")).strip()
        desig = str(row.get("Désignation", "")).strip()
        desig_lower = desig.lower()

        # Maj de la section courante + détection section vide
        if row_type == "section_header":
            current_section_code = code
            _skip_section = code not in active_sections
            if not _skip_section:
                if current_section_code not in section_articles:
                    section_articles[current_section_code] = []
                section_header_rows[current_section_code] = excel_row

        # Ignorer les lignes d'une section sans données (articles/recap internes)
        # Les recap_summary sont toujours affichés — même à 0 — pour un récapitulatif complet.
        if _skip_section and row_type in ("section_header", "sub_section", "article", "recap"):
            continue

        # ── BANDEAU "RÉCAPITULATIF" avant la première ligne recap_summary ──
        if row_type == "recap_summary" and not _recap_header_written:
            _recap_header_written = True
            # Ligne vide de séparation
            for c in range(1, max_col + 1):
                ws.cell(row=excel_row, column=c).fill = FILL_WHITE
            ws.row_dimensions[excel_row].height = 8
            excel_row += 1
            # Bandeau titre
            for c in range(1, max_col + 1):
                cell = ws.cell(row=excel_row, column=c)
                cell.fill = FILL_RECAP_HEADER
                cell.font = FONT_RECAP_HEADER
                cell.border = THIN_BORDER
            ws.cell(row=excel_row, column=1, value="RÉCAPITULATIF").alignment = Alignment(
                horizontal="left", vertical="center", indent=1
            )
            ws.row_dimensions[excel_row].height = 20
            excel_row += 1
        # ────────────────────────────────────────────────────────────────────────

        ws.cell(row=excel_row, column=1, value=code)
        ws.cell(row=excel_row, column=2, value=row["Désignation"])
        ws.cell(row=excel_row, column=3, value=row.get("Qu."))
        ws.cell(row=excel_row, column=4, value=row.get("U"))
        ws.cell(row=excel_row, column=5, value=row.get("Px_U_HT"))

        # --- Colonne F : TCO / Estimation ---
        # FIX CAS 2/3/4 : sub_section ET article contribuent au total de leur section.
        # Les sub_sections (Entete _Niv1/_Niv2) ont des prix propres (ex: 06.5.3.2)
        # et doivent apparaître dans le SUM recap — elles étaient silencieusement omises.
        if row_type in ("article", "sub_section") and current_section_code:
            section_articles[current_section_code].append(excel_row)

        if row_type == "article":
            ws.cell(row=excel_row, column=6, value=f"=C{excel_row}*E{excel_row}")

        elif row_type == "recap":
            # Total de la section (ligne grise)
            rows = section_articles.get(current_section_code, [])
            if rows:
                ws.cell(row=excel_row, column=6, value=_rows_to_sum_formula("F", rows))
            else:
                ws.cell(row=excel_row, column=6, value=0)
            section_total_row[current_section_code] = excel_row

        elif row_type == "recap_summary":
            # Ligne dans le tableau final récapitulatif
            recap_summary_rows.append(excel_row)
            # On cherche à lier au total de la section correspondante
            # On suppose que le code du récap correspond au code de la section
            target_row = section_total_row.get(code)
            if target_row:
                ws.cell(row=excel_row, column=6, value=f"=F{target_row}")
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))  # Fallback

        elif re.search(r"montant\s+ht", desig_lower):
            # Grand Total HT
            if recap_summary_rows:
                ws.cell(
                    row=excel_row, column=6, value=_rows_to_sum_formula("F", recap_summary_rows)
                )
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
            ht_row_idx = excel_row
        elif re.search(r"tva", desig_lower) and not re.search(r"ht", desig_lower):
            if ht_row_idx is not None:
                ws.cell(row=excel_row, column=6, value=f"=F{ht_row_idx}*{tva_rate}")
                tva_row_idx = excel_row
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
        elif re.search(r"montant\s+ttc", desig_lower):
            if ht_row_idx is not None and tva_row_idx is not None:
                ws.cell(row=excel_row, column=6, value=f"=F{ht_row_idx}+F{tva_row_idx}")
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
        elif row_type == "sub_section":
            # PARTIE 3 : formule =C*E si Qu. et Px_U_HT présents, sinon montant merger
            try:
                qu_val = float(row.get("Qu.", 0) or 0)
                px_u_val = float(row.get("Px_U_HT", 0) or 0)
            except (ValueError, TypeError):
                qu_val = px_u_val = 0.0
            if qu_val != 0 and px_u_val != 0:
                ws.cell(row=excel_row, column=6, value=f"=C{excel_row}*E{excel_row}")
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
        else:
            ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))

        # --- Colonnes Entreprises ---
        col_offset = 7
        for comp in companies:
            ws.cell(row=excel_row, column=col_offset, value=row.get(f"{comp}_Qu."))
            ws.cell(row=excel_row, column=col_offset + 1, value=row.get(f"{comp}_Px_U_HT"))

            qu_col = get_column_letter(col_offset)
            px_col = get_column_letter(col_offset + 1)
            tot_col = get_column_letter(col_offset + 2)

            if row_type == "article":
                ws.cell(
                    row=excel_row,
                    column=col_offset + 2,
                    value=f"={qu_col}{excel_row}*{px_col}{excel_row}",
                )
            elif row_type == "recap":
                rows = section_articles.get(current_section_code, [])
                if rows:
                    ws.cell(
                        row=excel_row,
                        column=col_offset + 2,
                        value=_rows_to_sum_formula(tot_col, rows),
                    )
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=0)
            elif row_type == "recap_summary":
                target_row = section_total_row.get(code)
                if target_row:
                    ws.cell(row=excel_row, column=col_offset + 2, value=f"={tot_col}{target_row}")
                else:
                    ws.cell(
                        row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT")
                    )
            elif re.search(r"montant\s+ht", desig_lower):
                if recap_summary_rows:
                    ws.cell(
                        row=excel_row,
                        column=col_offset + 2,
                        value=_rows_to_sum_formula(tot_col, recap_summary_rows),
                    )
                else:
                    ws.cell(
                        row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT")
                    )
            elif re.search(r"tva", desig_lower) and not re.search(r"ht", desig_lower):
                if ht_row_idx is not None:
                    ws.cell(
                        row=excel_row,
                        column=col_offset + 2,
                        value=f"={tot_col}{ht_row_idx}*{tva_rate}",
                    )
                else:
                    ws.cell(
                        row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT")
                    )
            elif re.search(r"montant\s+ttc", desig_lower):
                if ht_row_idx is not None and tva_row_idx is not None:
                    ws.cell(
                        row=excel_row,
                        column=col_offset + 2,
                        value=f"={tot_col}{ht_row_idx}+{tot_col}{tva_row_idx}",
                    )
                else:
                    ws.cell(
                        row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT")
                    )
            elif row_type == "sub_section":
                # PARTIE 3 : formule dynamique si données présentes
                try:
                    comp_qu = float(row.get(f"{comp}_Qu.", 0) or 0)
                    comp_px = float(row.get(f"{comp}_Px_U_HT", 0) or 0)
                except (ValueError, TypeError):
                    comp_qu = comp_px = 0.0
                if comp_qu != 0 and comp_px != 0:
                    ws.cell(
                        row=excel_row,
                        column=col_offset + 2,
                        value=f"={qu_col}{excel_row}*{px_col}{excel_row}",
                    )
                else:
                    ws.cell(
                        row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT")
                    )
            else:
                ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))

            ws.cell(row=excel_row, column=col_offset + 3, value=row.get(f"{comp}_Commentaire"))
            col_offset += 4

        # --- Format numérique (appliqué à toutes les lignes) ---
        ws.cell(row=excel_row, column=3).number_format = QTY_FORMAT  # Qu. TCO
        ws.cell(row=excel_row, column=5).number_format = MONEY_FORMAT  # Px U HT TCO
        ws.cell(row=excel_row, column=6).number_format = MONEY_FORMAT  # Px Tot HT TCO
        _ncol = 7
        for _ in companies:
            ws.cell(row=excel_row, column=_ncol).number_format = QTY_FORMAT  # Qu. entreprise
            ws.cell(
                row=excel_row, column=_ncol + 1
            ).number_format = MONEY_FORMAT  # Px U HT entreprise
            ws.cell(
                row=excel_row, column=_ncol + 2
            ).number_format = MONEY_FORMAT  # Px Tot HT entreprise
            _ncol += 4

        # --- Style ---
        font, fill = _get_row_style(row_type)

        # PARTIE 2 : sub_section sans prix → titre principal (BATIMENT F, etc.)
        if row_type == "sub_section":
            try:
                px_val = float(row.get("Px_Tot_HT", 0) or 0)
            except (ValueError, TypeError):
                px_val = 0.0
            if px_val == 0:
                font, fill = FONT_MAIN_TITLE, FILL_MAIN_TITLE

        # Hiérarchie visuelle : surcharge pour les lignes totaux généraux
        if row_type == "total_line":
            if re.search(r"montant\s+ht", desig_lower):
                font, fill = FONT_GRAND_TOTAL, FILL_MONTANT_HT
            elif re.search(r"\btva\b", desig_lower) and not re.search(r"\bht\b", desig_lower):
                font, fill = FONT_GRAND_TOTAL, FILL_TVA
            elif re.search(r"montant\s+ttc|(?<!\w)ttc(?!\w)", desig_lower):
                font, fill = FONT_GRAND_TOTAL, FILL_MONTANT_TTC

        # Sous-totaux intermédiaires (recap de section imbriquée) : 8pt comme la référence.
        # Un recap de section de niveau ≥ 2 (ex: "01.2.1") a 2 points dans son code
        # et reçoit une police plus petite que les totaux de section principale (11pt).
        if row_type == "recap" and code.count(".") >= 2:
            font = FONT_RECAP_SUB

        # Articles alternatifs : code suffixé par une ou plusieurs lettres après un chiffre
        # (ex: "01.1A", "10.2.2.1.4.1A", "11.2.1.1.2AA") → orange comme les modèles.
        if row_type in ("article", "sub_section", "section_header"):
            if re.search(r"\d[A-Z]+$", code):
                font = FONT_ALTERNATIVE

        _border = THIN_BORDER

        for c in range(1, max_col + 1):
            cell = ws.cell(row=excel_row, column=c)
            cell.font = font
            cell.border = _border
            # Toutes les cellules reçoivent un fill solide (même blanc) :
            # sans fill, les cellules transparentes laissent le texte de B déborder
            # sur les colonnes adjacentes pendant le scroll (freeze pane C3).
            cell.fill = fill if fill else FILL_WHITE
            # Alignement vertical uniforme sur toute la ligne → élimine le décalage
            # Code (bas par défaut) vs Désignation (top explicite).
            cell.alignment = Alignment(vertical="center")

        # Col A (Code) : left + center — même ancrage vertical que col B.
        ws.cell(row=excel_row, column=1).alignment = Alignment(horizontal="left", vertical="center")

        # Col B (Désignation) : wrap_text pour les textes longs + vertical center
        # pour rester aligné avec col A quelle que soit la hauteur de ligne.
        _indent = 0
        if row_type == "article":
            _indent = 2
        elif row_type == "sub_section" and font is not FONT_MAIN_TITLE:
            _indent = 1
        ws.cell(row=excel_row, column=2).alignment = Alignment(
            horizontal="left", indent=_indent, wrap_text=True, vertical="center"
        )

        if code and code in alert_by_code and row_type in ("article", "sub_section"):
            # Collecter la sévérité max + les entreprises concernées
            max_severity = "info"
            companies_with_errors: set[str] = set()
            companies_with_warnings: set[str] = set()
            has_global_error = False

            for alert in alert_by_code[code]:
                sev = alert["type"]
                comp = alert.get("company", "")
                if sev == "error":
                    max_severity = "error"
                    if comp:
                        companies_with_errors.add(comp)
                    else:
                        has_global_error = True
                elif sev == "warning" and max_severity != "error":
                    max_severity = "warning"
                    if comp:
                        companies_with_warnings.add(comp)

            if max_severity == "error":
                if has_global_error or not companies_with_errors:
                    # Erreur non liée à une entreprise → ligne entière
                    for c in range(1, max_col + 1):
                        ws.cell(row=excel_row, column=c).fill = FILL_ERROR
                else:
                    # Erreur ciblée → seulement les 4 colonnes de l'entreprise fautive
                    for comp in companies_with_errors:
                        if comp in companies:
                            _ci = companies.index(comp)
                            _sc = 7 + _ci * 4
                            for c in range(_sc, _sc + 4):
                                ws.cell(row=excel_row, column=c).fill = FILL_ERROR

            elif max_severity == "warning":
                if companies_with_warnings:
                    # Warning ciblé → 4 colonnes de l'entreprise concernée
                    for comp in companies_with_warnings:
                        if comp in companies:
                            _ci = companies.index(comp)
                            _sc = 7 + _ci * 4
                            for c in range(_sc, _sc + 4):
                                ws.cell(row=excel_row, column=c).fill = FILL_WARNING
                else:
                    # Fallback global : colonne Commentaire de toutes les entreprises
                    _c_off = 7
                    for _ in companies:
                        ws.cell(row=excel_row, column=_c_off + 3).fill = FILL_WARNING
                        _c_off += 4

        # Hauteur de ligne dynamique : s'adapte aux désignations longues qui wrappent.
        # Col B = 56.75 chars → ~55 chars utiles par ligne à 9pt.
        # Chaque ligne de texte ≈ 16 pt ; plancher 28.5 pt (1 ligne), plafond 80 pt.
        _n_lines = max(1, -(-len(desig) // 55)) if desig else 1
        ws.row_dimensions[excel_row].height = max(28.5, min(_n_lines * 16.0, 80.0))

        excel_row += 1

    # PARTIE 3 : Injection différée des formules pour section_header
    # Les plages d'articles/sub_sections ne sont connues qu'après le parcours complet.
    for sh_code, sh_row in section_header_rows.items():
        recap_row = section_total_row.get(sh_code)
        art_rows = section_articles.get(sh_code, [])

        # Col F : référencer le recap (=F{recap_row}) ou sommer les enfants
        if recap_row:
            ws.cell(row=sh_row, column=6, value=f"=F{recap_row}")
        elif art_rows:
            ws.cell(row=sh_row, column=6, value=_rows_to_sum_formula("F", art_rows))

        # Colonnes entreprises
        c_off = 7
        for _ in companies:
            tc = get_column_letter(c_off + 2)
            if recap_row:
                ws.cell(row=sh_row, column=c_off + 2, value=f"={tc}{recap_row}")
            elif art_rows:
                ws.cell(row=sh_row, column=c_off + 2, value=_rows_to_sum_formula(tc, art_rows))
            c_off += 4

    # Largeurs de colonnes — valeurs exactes de la référence
    ws.column_dimensions["A"].width = 9.5
    ws.column_dimensions["B"].width = 56.75
    ws.column_dimensions["C"].width = 9.5
    ws.column_dimensions["D"].width = 7.125
    ws.column_dimensions["E"].width = 14.125
    ws.column_dimensions["F"].width = 16.5
    for _ci in range(len(companies)):
        _cb = 7 + _ci * 4
        ws.column_dimensions[get_column_letter(_cb)].width = 9.5  # Qu.
        ws.column_dimensions[get_column_letter(_cb + 1)].width = 14.125  # Px U HT
        ws.column_dimensions[get_column_letter(_cb + 2)].width = 16.5  # Px Tot HT
        ws.column_dimensions[get_column_letter(_cb + 3)].width = 25.0  # Commentaire

    # Hauteurs en-têtes : 14.25 pt (conforme référence)
    ws.row_dimensions[header_row_1].height = 14.25
    ws.row_dimensions[header_row_2].height = 14.25

    # Freeze panes robuste + corrections anti-chevauchement
    fix_freeze_panes(ws, header_rows=header_row_2, frozen_cols=2)  # C{header_row_2 + 1} : lignes 1-header_row_2 + cols A-B
    fix_merged_cells_crossing_freeze(ws, header_rows=header_row_2, frozen_cols=2)  # retire fusions qui traversent
    prevent_text_overflow(ws, min_row=header_row_2 + 1, max_col=max_col)  # fill blanc sur cellules vides

    log.info("Workbook prêt. Output_path=%s", output_path)

    if output_path:
        wb.save(output_path)
        wb.close()
        return output_path
    else:
        buffer = io.BytesIO()
        wb.save(buffer)
        wb.close()
        buffer.seek(0)
        return buffer
