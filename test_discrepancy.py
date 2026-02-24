import pandas as pd
import os
import sys

# Add project root to path
sys.path.append(r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP")

from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from config import TVA_DEFAULT

tco_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 01 - DESAMIANTAGE - CURAGE - GROS OEUVRE.xlsx"
company_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx"
company_name = "GUSTAVE EIFFEL BDX"

print(f"--- Fichier TCO Modèle : {os.path.basename(tco_path)} ---")
try:
    tco_df, tco_meta = parse_tco(tco_path)
    print(f"TCO chargé : {len(tco_df)} lignes")
except Exception as e:
    print(f"Erreur chargement TCO : {e}")
    sys.exit(1)

print(f"\n--- Fichier Entreprise : {os.path.basename(company_path)} ---")
try:
    dpgf_df, dpgf_alerts = parse_dpgf(company_path)
    print(f"DPGF chargé : {len(dpgf_df)} lignes, {len(dpgf_alerts)} alertes")
    
    # Check DPGF totals
    art_dpgf = dpgf_df[dpgf_df["row_type"] == "article"]
    print(f"Nombre d'articles dans DPGF : {len(art_dpgf)}")
    print(f"Somme des Px_Tot_HT des articles dans DPGF : {art_dpgf['Px_Tot_HT'].sum():.2f}")
except Exception as e:
    print(f"Erreur chargement DPGF : {e}")
    sys.exit(1)

print(f"\n--- Fusion ---")
try:
    merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, company_name, tva_rate=0.20)
    print(f"Fusion terminée, {len(merge_alerts)} alertes de fusion")
    
    col_tot = f"{company_name}_Px_Tot_HT"
    
    total_ht = 0
    total_tva = 0
    total_ttc = 0
    
    for _, row in merged_df.iterrows():
        if row["row_type"] == "total_line":
            desig = str(row["Désignation"]).lower()
            val = row.get(col_tot)
            if "montant ht" in desig:
                total_ht = val
            elif "tva" in desig:
                total_tva = val
            elif "ttc" in desig:
                total_ttc = val
                
    print(f"\nRésultats calculés par l'app pour {company_name}:")
    print(f"MONTANT TOTAL HT  : {total_ht}")
    print(f"TVA               : {total_tva}")
    print(f"MONTANT TOTAL TTC : {total_ttc}")
    
    print("\n--- Comparaison avec les totaux lus dans le fichier entreprise (depuis diag_entreprise.py) ---")
    print("Fichier original HT : 249,389.00")
    print("Fichier original TVA: 49,877.80")
    print("Fichier original TTC: 299,266.80")
    
    if total_ht != 249389:
        print(f"\nDISCORDANCE HT : Écart de {total_ht - 249389:.2f} €")
        
        # Investigate missing articles
        dpgf_codes = set(dpgf_df[dpgf_df["row_type"] == "article"]["Code"].unique())
        tco_codes = set(tco_df[tco_df["row_type"] == "article"]["Code"].unique())
        
        missing_in_tco = dpgf_codes - tco_codes
        if missing_in_tco:
            print(f"Codes dans DPGF absents du TCO initial : {missing_in_tco}")
            
    # Check section totals summation
    section_totals = merged_df[merged_df["row_type"] == "section_header"][col_tot].sum()
    print(f"Somme des section_headers : {section_totals:.2f}")

except Exception as e:
    print(f"Erreur pendant la fusion : {e}")
    import traceback
    traceback.print_exc()
