import pandas as pd
from core.exporter import export_tco
from core.merger import merge_company_into_tco
import uuid

def test_tco_features():
    # Fake TCO Model
    tco_data = [
        {"Code": "01", "Désignation": "Chapitre 1", "row_type": "section_header"},
        {"Code": "01.1", "Désignation": "Article 1", "Qu.": 100, "U": "m2", "Px_U_HT": 10, "Px_Tot_HT": 1000, "row_type": "article"},
        {"Code": "02", "Désignation": "Chapitre 2", "row_type": "section_header"},
        {"Code": "02.1", "Désignation": "Article 2", "Qu.": 50, "U": "m3", "Px_U_HT": 20, "Px_Tot_HT": 1000, "row_type": "article"},
        {"Code": "", "Désignation": "MONTANT HT", "Px_Tot_HT": 2000, "row_type": "total_line"}
    ]
    tco_df = pd.DataFrame(tco_data)
    
    # Fake DPGF Offer with mismatches and extra line
    dpgf_data = [
        {"Code": "01.1", "Désignation": "Article 1", "Qu.": 1, "U": "Ens", "Px_U_HT": 1200, "Px_Tot_HT": 1200, "row_type": "article"}, # Mismatch Qty/Unit
        {"Code": "02.1", "Désignation": "Article 2", "Qu.": 50, "U": "m3", "Px_U_HT": 25, "Px_Tot_HT": 1250, "row_type": "article"}, # Normal
        {"Code": "02.2", "Désignation": "Ligne supplementaire", "Qu.": 2, "U": "u", "Px_U_HT": 100, "Px_Tot_HT": 200, "row_type": "article"} # Extra Line
    ]
    dpgf_df = pd.DataFrame(dpgf_data)

    print("Testing Merger...")
    merged, alerts = merge_company_into_tco(tco_df, dpgf_df, "ENTREPRISE_A")
    
    print("\nAlerts generated:")
    for a in alerts:
        print(f" - {a['message']}")
    
    print("\nDetecting is_extra_line:")
    if "is_extra_line" in merged.columns:
        extra = merged[merged["is_extra_line"] == True]
        print(extra[["Code", "Désignation"]])
    else:
        print("No extra lines column found.")

    print("\nTesting Exporter with metadata...")
    meta = {
        "project_info": {
            "moa": "Ville de Test",
            "moe": "ArchiTech",
            "devise": "€",
            "lot": "01 Gros Oeuvre"
        }
    }
    
    out_file = "test_export_output.xlsx"
    export_tco(merged, meta, output_path=out_file, alerts=alerts)
    print(f"Export OK: {out_file}")

if __name__ == "__main__":
    test_tco_features()
