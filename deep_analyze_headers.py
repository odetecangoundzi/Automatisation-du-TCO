import glob
import os

import pandas as pd


def deep_analyze_headers(directory):
    files = glob.glob(os.path.join(directory, "*.xls*"))
    report = []
    for f in files:
        try:
            xl = pd.ExcelFile(f)
            sheet_name = xl.sheet_names[0]
            df = pd.read_excel(f, sheet_name=sheet_name, header=None, nrows=100)

            found_header = False
            for i, row in df.iterrows():
                row_vals = [str(v).lower() for v in row.values if pd.notna(v)]
                # Check for "désignation" or "libellé" and "prix" or "total" or "unité"
                has_desig = any("signation" in v or "libellé" in v for v in row_vals)
                has_u = any(v == "u" or "unité" in v for v in row_vals)
                has_price = any(
                    "prix" in v or "p.u" in v or "p.u." in v or "montant" in v or "total" in v
                    for v in row_vals
                )

                if has_desig and (has_u or has_price):
                    headers = [str(v).strip() for v in row.values]
                    report.append(f"{os.path.basename(f)} (Row {i}): {' | '.join(headers)}")
                    found_header = True
                    break

            if not found_header:
                # List first few non-empty rows for debugging
                debug_rows = []
                for i, row in df.head(10).iterrows():
                    vals = [str(v).strip() for v in row.values if pd.notna(v)]
                    if vals:
                        debug_rows.append(f"Row {i}: {' | '.join(vals)}")
                report.append(
                    f"{os.path.basename(f)}: NO HEADER FOUND. First rows:\n" + "\n".join(debug_rows)
                )

        except Exception as e:
            report.append(f"{os.path.basename(f)}: ERROR {str(e)}")

    with open("deep_header_analysis.txt", "w", encoding="utf-8") as out:
        out.write("\n\n" + "=" * 50 + "\n\n".join(report))


deep_analyze_headers(r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\Autres_Formats")
