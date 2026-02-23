import os
import traceback
from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco

ENTREPRISE_DIR = "entreprise"
OUTPUT_FILE = "merge_results_utf8.txt"

TCO_FILES = {
    "01": "DPGF LOT 01 - DESAMIANTAGE - CURAGE - GROS OEUVRE.xlsx",
    "02": "DPGF LOT 02 - MENUISERIES - SERRURERIE.xlsx",
    "03": "DPGF LOT 03 - PLATRERIE.xlsx",
    "04": "DPGF LOT 04 - REVETEMENT INTERIEUR ET EXTERIEUR - PEINTURE - SIGNALETIQUE.xlsx",
    "05": "DPGF LOT 05 - PLOMBERIE - CVC.xlsx",
    "06": "DPGF LOT 06 - ELECTRICITE.xlsx",
}

files = sorted(os.listdir(ENTREPRISE_DIR))

with open(OUTPUT_FILE, "w", encoding="utf-8") as out:
    for fname in files:
        if not fname.endswith('.xlsx'):
            continue
        
        lot = None
        if "LOT 01" in fname.upper(): lot = "01"
        elif "LOT 02" in fname.upper(): lot = "02"
        elif "LOT 03" in fname.upper(): lot = "03"
        elif "LOT 04" in fname.upper(): lot = "04"
        elif "LOT 05" in fname.upper(): lot = "05"
        elif "LOT 06" in fname.upper() or "EIFFEL BDX" in fname.upper(): lot = "06"
        
        out.write(f"{'='*80}\n")
        out.write(f"FICHIER : {fname}\n")
        if not lot:
            out.write("  -> Impossible de déduire le LOT.\n")
            continue
            
        tco_file = TCO_FILES.get(lot)
        if not tco_file or not os.path.exists(tco_file):
            out.write(f"  -> TCO introuvable pour le LOT {lot}: {tco_file}\n")
            continue
            
        try:
            tco_df, tco_meta = parse_tco(tco_file)
            dpgf_df, dpgf_alerts = parse_dpgf(os.path.join(ENTREPRISE_DIR, fname))
            
            merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, "TEST", tva_rate=20.0)
            
            col_total = "TEST_Px_Tot_HT"
            if col_total in merged_df.columns:
                matched = merged_df[merged_df[col_total] > 0]
                total_sum = merged_df[col_total].sum()
            else:
                matched = []
                total_sum = 0.0
            
            dpgf_with_price = len(dpgf_df[(dpgf_df['row_type'] == 'article') & (dpgf_df['Px_Tot_HT'] > 0)])
            
            out.write(f"  Fusion avec {tco_file}...\n")
            out.write(f"  Articles avec prix > 0 dans le DPGF    : {dpgf_with_price}\n")
            out.write(f"  Articles avec prix > 0 dans le FUSIONNÉ: {len(matched)}\n")
            out.write(f"  Somme totale fusionnée                 : {total_sum:.2f} €\n")
            
            if len(matched) < dpgf_with_price * 0.8:
                out.write(f"  ⚠️ FAIBLE TAUX DE CORRESPONDANCE ! ({len(matched)} / {dpgf_with_price})\n")
                
                # Show top unmatched
                dpgf_codes = set(dpgf_df[(dpgf_df['row_type'] == 'article') & (dpgf_df['Px_Tot_HT'] > 0)]['Code'].astype(str))
                tco_codes = set(tco_df[tco_df['row_type'] == 'article']['Code'].astype(str))
                unmatched = sorted(list(dpgf_codes - tco_codes))
                if unmatched:
                    out.write(f"  -> 5 premiers codes du DPGF non trouvés dans le TCO: {unmatched[:5]}\n")
                
        except Exception as e:
            out.write(f"  ❌ ERREUR LORS DE LA FUSION : {e}\n")
            out.write(traceback.format_exc())

    out.write(f"{'='*80}\n")
print(f"Results written to {OUTPUT_FILE}")
