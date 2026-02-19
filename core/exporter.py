"""
exporter.py — Export du TCO fusionné en fichier Excel formaté.

Génère un fichier .xlsx avec :
- Headers groupés par entreprise (merged cells)
- Mise en forme (gras, freeze pane, largeur auto)
- Coloration selon alertes
- Colonnes masquées (Entete, flags internes)
"""

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import pandas as pd


# ---------------------------------------------------------------------------
# Constantes de style
# ---------------------------------------------------------------------------

FONT_HEADER = Font(bold=True, size=11, color="FFFFFF")
FONT_HEADER_COMPANY = Font(bold=True, size=12, color="FFFFFF")
FONT_TOTAL = Font(bold=True, size=10)
FONT_DATA = Font(size=10)

FILL_HEADER = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
FILL_COMPANY_COLORS = [
    PatternFill(start_color="548235", end_color="548235", fill_type="solid"),  # Vert
    PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid"),  # Doré
    PatternFill(start_color="843C0C", end_color="843C0C", fill_type="solid"),  # Marron
    PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid"),  # Violet
    PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"),  # Rouge
]
FILL_TOTAL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
FILL_SECTION = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")

# Couleurs d'alertes
FILL_ERROR = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_WARNING = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
FILL_NOTE = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
FILL_INFO = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"),
    right=Side(style="thin"),
    top=Side(style="thin"),
    bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _detect_companies(df):
    """
    Détecte les noms d'entreprises à partir des colonnes du DataFrame.
    Retourne une liste de noms uniques d'entreprises.
    """
    companies = []
    seen = set()
    for col in df.columns:
        if col.endswith("_Qu."):
            name = col[:-4]  # enlève "_Qu."
            if name not in seen:
                companies.append(name)
                seen.add(name)
    return companies


def _auto_width(ws, min_width=8, max_width=40):
    """Ajuste la largeur des colonnes automatiquement."""
    for col_cells in ws.columns:
        col_letter = get_column_letter(col_cells[0].column)
        max_len = min_width
        for cell in col_cells:
            if cell.value:
                cell_len = len(str(cell.value))
                max_len = max(max_len, min(cell_len + 2, max_width))
        ws.column_dimensions[col_letter].width = max_len


# ---------------------------------------------------------------------------
# Main exporter
# ---------------------------------------------------------------------------

def export_tco(merged_df, meta, output_path, alerts=None):
    """
    Exporte le TCO fusionné en fichier Excel formaté.

    Args:
        merged_df   : DataFrame fusionné
        meta        : dict de métadonnées (de parse_tco)
        output_path : chemin du fichier de sortie
        alerts      : liste d'alertes optionnelle (pour la coloration)
    """
    if alerts is None:
        alerts = []

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = meta.get("sheet_name", "TCO Final")

    companies = _detect_companies(merged_df)

    # --- Construire les colonnes de sortie ---
    base_cols = ["Code", "Désignation", "Qu.", "U", "Px_U_HT", "Px_Tot_HT"]
    company_cols = {}
    for comp in companies:
        company_cols[comp] = [
            f"{comp}_Qu.",
            f"{comp}_Px_U_HT",
            f"{comp}_Px_Tot_HT",
            f"{comp}_Commentaire",
        ]

    # --- ROW 1 : Headers groupés ---
    current_col = 1

    # Estimation header (colonnes C-F)
    ws.cell(row=1, column=2, value="Etudes")
    ws.cell(row=1, column=3, value=" Estimation")
    ws.merge_cells(
        start_row=1, start_column=3,
        end_row=1, end_column=6
    )
    for c in range(2, 7):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")

    # Company headers
    company_start_col = 7  # après F (col 6)
    for comp_idx, comp in enumerate(companies):
        start_col = company_start_col
        end_col = start_col + 3  # 4 colonnes par entreprise

        ws.cell(row=1, column=start_col, value=comp)
        ws.merge_cells(
            start_row=1, start_column=start_col,
            end_row=1, end_column=end_col
        )

        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = FONT_HEADER_COMPANY
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")

        company_start_col = end_col + 1

    # --- ROW 2 : Noms de colonnes ---
    col_headers_row2 = ["Code", "Désignation", "Qu.", "U", "Px U. HT", "Px Tot HT"]
    for i, header in enumerate(col_headers_row2, 1):
        cell = ws.cell(row=2, column=i, value=header)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER

    col_offset = 7
    for comp_idx, comp in enumerate(companies):
        comp_headers = ["Qu.", "Px U. HT", "Px Tot HT", "Commentaire poste"]
        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for i, header in enumerate(comp_headers):
            cell = ws.cell(row=2, column=col_offset + i, value=header)
            cell.font = FONT_HEADER
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")
            cell.border = THIN_BORDER
        col_offset += 4

    # --- Construire l'index des alertes par code ---
    alert_by_code = {}
    for alert in alerts:
        code = alert.get("code", "")
        if code:
            if code not in alert_by_code:
                alert_by_code[code] = []
            alert_by_code[code].append(alert)

    # --- ROWS 3+ : Données ---
    for df_idx, row in merged_df.iterrows():
        excel_row = df_idx + 3  # row 3 = première ligne de données

        code = row["Code"]
        row_type = row["row_type"]

        # Colonnes de base
        ws.cell(row=excel_row, column=1, value=code)
        ws.cell(row=excel_row, column=2, value=row["Désignation"])
        ws.cell(row=excel_row, column=3, value=row.get("Qu."))
        ws.cell(row=excel_row, column=4, value=row.get("U"))
        ws.cell(row=excel_row, column=5, value=row.get("Px_U_HT"))
        ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))

        # Colonnes entreprises
        col_offset = 7
        for comp in companies:
            ws.cell(
                row=excel_row, column=col_offset,
                value=row.get(f"{comp}_Qu.")
            )
            ws.cell(
                row=excel_row, column=col_offset + 1,
                value=row.get(f"{comp}_Px_U_HT")
            )
            ws.cell(
                row=excel_row, column=col_offset + 2,
                value=row.get(f"{comp}_Px_Tot_HT")
            )
            ws.cell(
                row=excel_row, column=col_offset + 3,
                value=row.get(f"{comp}_Commentaire")
            )
            col_offset += 4

        # --- Mise en forme selon le type de ligne ---
        max_col = 6 + len(companies) * 4

        if row_type == "total":
            for c in range(1, max_col + 1):
                cell = ws.cell(row=excel_row, column=c)
                cell.font = FONT_TOTAL
                cell.fill = FILL_TOTAL
                cell.border = THIN_BORDER
            # Merger A:B pour les totaux
            ws.merge_cells(
                start_row=excel_row, start_column=1,
                end_row=excel_row, end_column=2
            )
        elif row_type in ("data", "other", "recap"):
            # Style normal
            for c in range(1, max_col + 1):
                cell = ws.cell(row=excel_row, column=c)
                cell.font = FONT_DATA
                cell.border = THIN_BORDER

            # Coloration des alertes par code
            if code and code in alert_by_code:
                for alert in alert_by_code[code]:
                    fill = _get_alert_fill(alert["color"])
                    if fill:
                        # Colorer les colonnes entreprise concernées
                        comp_col_start = 7
                        for comp in companies:
                            for c in range(comp_col_start, comp_col_start + 4):
                                ws.cell(row=excel_row, column=c).fill = fill
                            comp_col_start += 4

    # --- Freeze pane ---
    ws.freeze_panes = "C3"

    # --- Largeur auto ---
    _auto_width(ws)

    # Largeur fixe pour Désignation (plus large)
    ws.column_dimensions["B"].width = 50

    # --- Sauvegarde ---
    wb.save(output_path)
    wb.close()

    return output_path


def _get_alert_fill(color):
    """Retourne le PatternFill correspondant à la couleur d'alerte."""
    mapping = {
        "red": FILL_ERROR,
        "orange": FILL_WARNING,
        "yellow": FILL_NOTE,
        "blue": FILL_INFO,
    }
    return mapping.get(color)
