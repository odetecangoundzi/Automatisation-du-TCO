import glob
import os

import pandas as pd


def analyze_headers(directory):
    files = glob.glob(os.path.join(directory, "*.xls*"))
    report = []
    for f in files:
        try:
            # Try to read the first few rows to find the header
            df = pd.read_excel(f, header=None, nrows=20)
            # Simple heuristic to find header row (like in the app)
            header_row = -1
            for i, row in df.iterrows():
                row_vals = [str(v).lower() for v in row.values]
                if any("code" in v for v in row_vals) and any(
                    "signation" in v or "libellé" in v for v in row_vals
                ):
                    header_row = i
                    break

            if header_row != -1:
                headers = [str(v).strip() for v in df.iloc[header_row].values]
                report.append(f"{os.path.basename(f)}: {' | '.join(headers)}")
            else:
                report.append(f"{os.path.basename(f)}: NO HEADER FOUND in first 20 rows")
        except Exception as e:
            report.append(f"{os.path.basename(f)}: ERROR {str(e)}")

    with open("header_analysis.txt", "w", encoding="utf-8") as out:
        out.write("\n".join(report))


analyze_headers(r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\Autres_Formats")
