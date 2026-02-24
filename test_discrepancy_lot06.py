import pandas as pd
import os
import sys

# Add project root to path
sys.path.append(r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP")

from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from config import TVA_DEFAULT

tco_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 06 - ELECTRICITE.xlsx"
company_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx"
company_name = "GUSTAVE EIFFEL BDX"

print(f"--- Fichier TCO Modèle : {os.path.basename(tco_path)} ---")
tco_df, tco_meta = parse_tco(tco_path)
print(f"TCO chargé : {len(tco_df)} lignes")

print(f"\n--- Fichier Entreprise : {os.path.basename(company_path)} ---")
dpgf_df, dpgf_alerts = parse_dpgf(company_path)
print(f"DPGF chargé : {len(dpgf_df)} lignes")

print(f"\n--- Fusion ---")
merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, company_name, tva_rate=0.20)

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
print(f"MONTANT TOTAL HT  : {total_ht:,.2f}")
print(f"TVA               : {total_tva:,.2f}")
print(f"MONTANT TOTAL TTC : {total_ttc:,.2f}")

print("\n--- Alertes de fusion ---")
for a in merge_alerts:
    print(f"[{a['type'].upper()}] {a['message']}")

print("\n--- Comparaison avec les totaux attendus ---")
print("Fichier original HT : 249,389.00")
print("Fichier original TVA: 49,877.80")
print("Fichier original TTC: 299,266.80")

# Check for Decimal types
try:
    from decimal import Decimal
    if isinstance(total_ht, Decimal):
        print("\n✅ Montants calculés en type Decimal.")
    else:
        print(f"\n❌ Montants non calculés en type Decimal (type={type(total_ht)}).")
except:
    pass

# Identify missing or extra items
dpgf_art = dpgf_df[dpgf_df["row_type"] == "article"]
tco_art = tco_df[tco_df["row_type"] == "article"]

dpgf_codes = set(dpgf_art["Code"].dropna().unique())
tco_codes = set(tco_art["Code"].dropna().unique())

print(f"\nCodes dans DPGF mais absents du TCO : {dpgf_codes - tco_codes}")
print(f"Codes dans TCO mais absents du DPGF : {tco_codes - dpgf_codes}")

# Check for dynamic insertions
dynamic_inserted = [a["code"] for a in merge_alerts if "Insertion dynamique" in a.get("message", "")]
print(f"\nInsertions dynamiques effectuées : {dynamic_inserted}")

# Check if some codes in DPGF have NO prices but are classified as articles
print("\nArticles du DPGF avec Px_Tot_HT == 0:")
zero_prices = dpgf_art[dpgf_art["Px_Tot_HT"] == 0][["Code", "Désignation"]]
print(zero_prices.to_string())
