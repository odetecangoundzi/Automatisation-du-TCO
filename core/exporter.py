"""
exporter.py — Export du TCO fusionné en fichier Excel formaté.

Génère un fichier .xlsx avec :
- Headers groupés par entreprise (merged cells)
- Mise en forme (gras, freeze pane, largeur auto)
- Coloration selon alertes
- Lignes section_header et recap mises en évidence
- Support export via BytesIO (pas de sauvegarde disque obligatoire)
"""

import io
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Constantes de style
# ---------------------------------------------------------------------------

FONT_HEADER         = Font(bold=True, size=11, color="FFFFFF")
FONT_HEADER_COMPANY = Font(bold=True, size=12, color="FFFFFF")
FONT_SECTION        = Font(bold=True, size=10, color="1F4E79")
FONT_RECAP          = Font(bold=True, size=10, color="2F5496")
FONT_TOTAL          = Font(bold=True, size=10)
FONT_DATA           = Font(size=10)

FILL_HEADER = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
FILL_COMPANY_COLORS = [
    PatternFill(start_color="548235", end_color="548235", fill_type="solid"),
    PatternFill(start_color="BF8F00", end_color="BF8F00", fill_type="solid"),
    PatternFill(start_color="843C0C", end_color="843C0C", fill_type="solid"),
    PatternFill(start_color="7030A0", end_color="7030A0", fill_type="solid"),
    PatternFill(start_color="C00000", end_color="C00000", fill_type="solid"),
]

FILL_SECTION       = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
FILL_RECAP         = PatternFill(start_color="B4C6E7", end_color="B4C6E7", fill_type="solid")
FILL_RECAP_SUMMARY = PatternFill(start_color="D6DCE4", end_color="D6DCE4", fill_type="solid")
FILL_TOTAL_LINE    = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
FILL_SUB_SECTION   = PatternFill(start_color="EDF2F9", end_color="EDF2F9", fill_type="solid")

FILL_ERROR   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
FILL_WARNING = PatternFill(start_color="FFE4B5", end_color="FFE4B5", fill_type="solid")
FILL_NOTE    = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid")
FILL_INFO    = PatternFill(start_color="D6EAF8", end_color="D6EAF8", fill_type="solid")

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"),  bottom=Side(style="thin"),
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _detect_companies(df):
    """Détecte les noms d'entreprises à partir des colonnes _Qu."""
    companies, seen = [], set()
    for col in df.columns:
        if col.endswith("_Qu."):
            name = col[:-4]
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
                # Truncate string representation to check length, avoid huge cells
                val_str = str(cell.value)
                line_len = max(len(line) for line in val_str.split('\n')) if '\n' in val_str else len(val_str)
                max_len = max(max_len, min(line_len + 2, max_width))
        ws.column_dimensions[col_letter].width = max_len


def _get_alert_fill(color):
    return {"red": FILL_ERROR, "orange": FILL_WARNING,
            "yellow": FILL_NOTE, "blue": FILL_INFO}.get(color)


def _get_row_style(row_type):
    return {
        "section_header": (FONT_SECTION, FILL_SECTION),
        "recap":          (FONT_RECAP,   FILL_RECAP),
        "recap_summary":  (FONT_RECAP,   FILL_RECAP_SUMMARY),
        "total_line":     (FONT_TOTAL,   FILL_TOTAL_LINE),
        "sub_section":    (FONT_DATA,    FILL_SUB_SECTION),
    }.get(row_type, (FONT_DATA, None))


# ---------------------------------------------------------------------------
# Main exporter
# ---------------------------------------------------------------------------

def export_tco(merged_df, meta, output_path=None, alerts=None):
    """
    Exporte le TCO fusionné en fichier Excel formaté.
    """
    if alerts is None:
        alerts = []

    log.info("Début export Excel. Lignes=%d", len(merged_df))
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = meta.get("sheet_name", "TCO Final")

    companies = _detect_companies(merged_df)
    log.debug("Entreprises détectées : %s", companies)

    # --- ROW 1 : Headers groupés ---
    ws.cell(row=1, column=2, value="Etudes")
    ws.cell(row=1, column=3, value=" Estimation")
    ws.merge_cells(start_row=1, start_column=3, end_row=1, end_column=6)
    for c in range(2, 7):
        cell = ws.cell(row=1, column=c)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")

    company_start_col = 7
    for comp_idx, comp in enumerate(companies):
        start_col = company_start_col
        end_col   = start_col + 3
        ws.cell(row=1, column=start_col, value=comp)
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=end_col)
        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for c in range(start_col, end_col + 1):
            cell = ws.cell(row=1, column=c)
            cell.font = FONT_HEADER_COMPANY
            cell.fill = fill
            cell.alignment = Alignment(horizontal="center")
        company_start_col = end_col + 1

    # --- ROW 2 : Noms de colonnes ---
    base_headers = ["Code", "Désignation", "Qu.", "U", "Px U. HT", "Px Tot HT"]
    for i, header in enumerate(base_headers, 1):
        cell = ws.cell(row=2, column=i, value=header)
        cell.font = FONT_HEADER
        cell.fill = FILL_HEADER
        cell.alignment = Alignment(horizontal="center")
        cell.border = THIN_BORDER

    col_offset = 7
    for comp_idx, comp in enumerate(companies):
        fill = FILL_COMPANY_COLORS[comp_idx % len(FILL_COMPANY_COLORS)]
        for i, header in enumerate(["Qu.", "Px U. HT", "Px Tot HT", "Commentaire poste"]):
            cell = ws.cell(row=2, column=col_offset + i, value=header)
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

    # --- ROWS 3+ : Données ---
    excel_row = 3
    
    # Tracking pour les formules dynamiques
    section_articles  = {} # { '01.1': [row_idx, ...], ... }
    section_total_row = {} # { '01.1': row_idx, ... }
    recap_summary_rows = [] # [row_idx, ...]
    
    current_section_code = None
    
    # On fait un premier passage pour identifier les lignes et types si nécessaire ?
    # Non, on peut faire en un passage car les articles précèdent leurs totaux,
    # et les totaux précèdent le récap (généralement).
    
    for _, row in merged_df.iterrows():
        row_type = row["row_type"]
        if row_type == "empty":
            continue

        code  = str(row.get("Code", "")).strip()
        desig = str(row.get("Désignation", "")).strip().lower()

        # Maj de la section courante
        if row_type == "section_header":
            current_section_code = code
            if current_section_code not in section_articles:
                section_articles[current_section_code] = []

        ws.cell(row=excel_row, column=1, value=code)
        ws.cell(row=excel_row, column=2, value=row["Désignation"])
        ws.cell(row=excel_row, column=3, value=row.get("Qu."))
        ws.cell(row=excel_row, column=4, value=row.get("U"))
        ws.cell(row=excel_row, column=5, value=row.get("Px_U_HT"))
        
        # --- Colonne F : TCO / Estimation ---
        if row_type == "article":
            if current_section_code:
                section_articles[current_section_code].append(excel_row)
            ws.cell(row=excel_row, column=6, value=f"=C{excel_row}*E{excel_row}")
            
        elif row_type == "recap":
            # Total de la section (ligne grise)
            rows = section_articles.get(current_section_code, [])
            if rows:
                formula = "=SUM(" + ",".join([f"F{r}" for r in rows]) + ")"
                ws.cell(row=excel_row, column=6, value=formula)
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
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT")) # Fallback
                
        elif "montant ht" in desig:
            # Grand Total HT
            if recap_summary_rows:
                formula = "=SUM(" + ",".join([f"F{r}" for r in recap_summary_rows]) + ")"
                ws.cell(row=excel_row, column=6, value=formula)
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
            ht_row_idx = excel_row
            
        elif "tva" in desig and "ht" not in desig:
            if 'ht_row_idx' in locals() and ht_row_idx:
                ws.cell(row=excel_row, column=6, value=f"=F{ht_row_idx}*0.2")
                tva_row_idx = excel_row
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
                
        elif "montant ttc" in desig:
            if 'ht_row_idx' in locals() and 'tva_row_idx' in locals() and ht_row_idx and tva_row_idx:
                ws.cell(row=excel_row, column=6, value=f"=F{ht_row_idx}+F{tva_row_idx}")
            else:
                ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))
        else:
            ws.cell(row=excel_row, column=6, value=row.get("Px_Tot_HT"))

        # --- Colonnes Entreprises ---
        col_offset = 7
        for comp in companies:
            ws.cell(row=excel_row, column=col_offset,     value=row.get(f"{comp}_Qu."))
            ws.cell(row=excel_row, column=col_offset + 1, value=row.get(f"{comp}_Px_U_HT"))
            
            qu_col  = get_column_letter(col_offset)
            px_col  = get_column_letter(col_offset + 1)
            tot_col = get_column_letter(col_offset + 2)
            
            if row_type == "article":
                ws.cell(row=excel_row, column=col_offset + 2, value=f"={qu_col}{excel_row}*{px_col}{excel_row}")
            elif row_type == "recap":
                rows = section_articles.get(current_section_code, [])
                if rows:
                    formula = f"=SUM(" + ",".join([f"{tot_col}{r}" for r in rows]) + ")"
                    ws.cell(row=excel_row, column=col_offset + 2, value=formula)
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=0)
            elif row_type == "recap_summary":
                target_row = section_total_row.get(code)
                if target_row:
                    ws.cell(row=excel_row, column=col_offset + 2, value=f"={tot_col}{target_row}")
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))
            elif "montant ht" in desig:
                if recap_summary_rows:
                    formula = f"=SUM(" + ",".join([f"{tot_col}{r}" for r in recap_summary_rows]) + ")"
                    ws.cell(row=excel_row, column=col_offset + 2, value=formula)
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))
            elif "tva" in desig and "ht" not in desig:
                if 'ht_row_idx' in locals() and ht_row_idx:
                    ws.cell(row=excel_row, column=col_offset + 2, value=f"={tot_col}{ht_row_idx}*0.2")
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))
            elif "montant ttc" in desig:
                if 'ht_row_idx' in locals() and 'tva_row_idx' in locals() and ht_row_idx and tva_row_idx:
                    ws.cell(row=excel_row, column=col_offset + 2, value=f"={tot_col}{ht_row_idx}+{tot_col}{tva_row_idx}")
                else:
                    ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))
            else:
                ws.cell(row=excel_row, column=col_offset + 2, value=row.get(f"{comp}_Px_Tot_HT"))
                
            ws.cell(row=excel_row, column=col_offset + 3, value=row.get(f"{comp}_Commentaire"))
            col_offset += 4

        # Style ...
        font, fill = _get_row_style(row_type)
        for c in range(1, max_col + 1):
            cell = ws.cell(row=excel_row, column=c)
            cell.font = font
            cell.border = THIN_BORDER
            if fill:
                cell.fill = fill

        if code and code in alert_by_code and row_type == "article":
            for alert in alert_by_code[code]:
                alert_fill = _get_alert_fill(alert["color"])
                if alert_fill:
                    comp_col = 7
                    for _ in companies:
                        for c in range(comp_col, comp_col + 4):
                            ws.cell(row=excel_row, column=c).fill = alert_fill
                        comp_col += 4

        excel_row += 1

    ws.freeze_panes = "C3"
    _auto_width(ws)
    ws.column_dimensions["B"].width = 55 # Force Désignation width

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
