import sys
sys.path.insert(0, r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP")
import pandas as pd
from decimal import Decimal
from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco
from core.utils import classify_row

print("=" * 70)
print("DIAGNOSTIC CAS 1 : LOT 04 - codes manquants mais montants presents")
print("=" * 70)

dpgf04_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\DPGF LOT 04 - REVETEMENT INTERIEUR ET EXTERIEUR - PEINTURE - SIGNALETIQUE.xlsx"
tco04_path  = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 04 - REVETEMENT INTERIEUR ET EXTERIEUR - PEINTURE - SIGNALETIQUE.xlsx"

dpgf04_df, dpgf04_alerts = parse_dpgf(dpgf04_path)
print(f"DPGF LOT 04 entreprise: {len(dpgf04_df)} lignes, {len(dpgf04_alerts)} alertes")

# Lignes sans code mais avec montants
no_code = dpgf04_df[
    (dpgf04_df["Code"].str.strip() == "") &
    (dpgf04_df["Px_Tot_HT"] != Decimal("0.0"))
]
print(f"\nLignes sans Code mais Px_Tot_HT != 0: {len(no_code)}")
for _, r in no_code.iterrows():
    print(f"  row={r['original_row']} | row_type={r['row_type']} | Px_Tot={r['Px_Tot_HT']} | Desig={r['Désignation'][:60]} | Entete={repr(r['Entete'])}")

# Lignes sans code sans montants (titres, etc.)
no_code_no_price = dpgf04_df[
    (dpgf04_df["Code"].str.strip() == "") &
    (dpgf04_df["Px_Tot_HT"] == Decimal("0.0"))
]
print(f"\nLignes sans Code et sans montant: {len(no_code_no_price)}")
for _, r in no_code_no_price.head(10).iterrows():
    print(f"  row={r['original_row']} | row_type={r['row_type']} | Desig={r['Désignation'][:60]} | Entete={repr(r['Entete'])}")

# Distribution des row_types
print(f"\nDistribution row_types DPGF LOT 04:")
print(dpgf04_df["row_type"].value_counts().to_string())

# Les 20 premières lignes classifiées
print(f"\nPremières 25 lignes du DPGF LOT 04 (code, type, entete, tot):")
for _, r in dpgf04_df.head(25).iterrows():
    print(f"  Code={r['Code'][:20]:20s} | type={r['row_type']:15s} | Entete={repr(r['Entete'])[:30]:30s} | Px_Tot={r['Px_Tot_HT']} | Desig={r['Désignation'][:40]}")

# TCO LOT 04
try:
    tco04_df, tco04_meta = parse_tco(tco04_path)
    print(f"\nTCO LOT 04 template: {len(tco04_df)} lignes")
    
    # Merge
    merged04, alerts04 = merge_company_into_tco(tco04_df, dpgf04_df, "LOT04_ENT", tva_rate=0.20)
    col04 = "LOT04_ENT_Px_Tot_HT"
    print(f"Alertes fusion LOT 04: {len(alerts04)}")
    for a in alerts04[:10]:
        print(f"  [{a['type'].upper()}] {a.get('code','')} : {a['message']}")
    
    print(f"\nTotaux finaux LOT 04:")
    for _, r in merged04[merged04["row_type"] == "total_line"].iterrows():
        print(f"  {r['Désignation'][:50]} : {r.get(col04)}")
    
    # Sum all articles
    art04 = dpgf04_df[dpgf04_df["row_type"] == "article"]
    print(f"\nSomme articles DPGF LOT 04: {art04['Px_Tot_HT'].apply(lambda x: float(x) if isinstance(x, Decimal) else float(x) if x is not None else 0).sum():.2f}")
    
except Exception as e:
    print(f"Erreur TCO LOT 04: {e}")
    import traceback; traceback.print_exc()

print("\n" + "=" * 70)
print("DIAGNOSTIC CAS 2 : LOT 02 - 02.5.3.2 absent du total SALLE POLYVALENTE")
print("=" * 70)

dpgf02_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\14-DE-20251282 - DPGF LOT 02 - MENUISERIES - SERRURERIE - APPS MUSCULATION.xlsx"
tco02_path  = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 02 - MENUISERIES - SERRURERIE.xlsx"

dpgf02_df, dpgf02_alerts = parse_dpgf(dpgf02_path)
print(f"DPGF LOT 02 entreprise: {len(dpgf02_df)} lignes, {len(dpgf02_alerts)} alertes")

# Chercher 02.5.3.2
code_target = "02.5.3.2"
rows = dpgf02_df[dpgf02_df["Code"] == code_target]
print(f"\n{code_target} dans DPGF:")
if rows.empty:
    print(f"  ABSENT du DPGF")
    # Try partial match
    partial = dpgf02_df[dpgf02_df["Code"].str.contains("5.3.2", na=False)]
    print(f"  Recherche partielle '5.3.2': {len(partial)} résultats")
    for _, r in partial.iterrows():
        print(f"    Code={r['Code']} | row_type={r['row_type']} | Px_Tot={r['Px_Tot_HT']} | Desig={r['Désignation'][:50]}")
else:
    for _, r in rows.iterrows():
        print(f"  Code={r['Code']} | row_type={r['row_type']} | Entete={repr(r['Entete'])} | Qu={r['Qu.']} | PU={r['Px_U_HT']} | Tot={r['Px_Tot_HT']}")

# Section 02.5 context
section_025 = dpgf02_df[dpgf02_df["Code"].str.startswith("02.5", na=False)]
print(f"\nToutes les lignes 02.5.* dans le DPGF:")
for _, r in section_025.iterrows():
    print(f"  Code={r['Code']:20s} | type={r['row_type']:15s} | Entete={repr(r['Entete'])[:30]:30s} | Tot={r['Px_Tot_HT']} | Desig={r['Désignation'][:50]}")

# SALLE POLYVALENTE recherche
salle_poly = dpgf02_df[dpgf02_df["Désignation"].str.contains("POLYVALENTE|POLIVAL", case=False, na=False)]
print(f"\nLignes 'SALLE POLYVALENTE' dans DPGF: {len(salle_poly)}")
for _, r in salle_poly.iterrows():
    print(f"  Code={r['Code']} | type={r['row_type']} | Tot={r['Px_Tot_HT']} | Entete={repr(r['Entete'])}")

# TCO LOT 02
try:
    tco02_df, _ = parse_tco(tco02_path)
    print(f"\nTCO LOT 02: {len(tco02_df)} lignes")
    
    # Check 02.5.3.2 in TCO
    in_tco = tco02_df[tco02_df["Code"] == code_target]
    print(f"\n{code_target} dans TCO template: {'PRESENT' if not in_tco.empty else 'ABSENT'}")
    if not in_tco.empty:
        r = in_tco.iloc[0]
        print(f"  row_type={r['row_type']} | Entete={repr(r['Entete'])}")
    
    # Merge
    merged02, alerts02 = merge_company_into_tco(tco02_df, dpgf02_df, "LOT02_ENT", tva_rate=0.20)
    col02 = "LOT02_ENT_Px_Tot_HT"
    print(f"\nAlertes fusion: {len(alerts02)}")
    for a in alerts02[:15]:
        print(f"  [{a['type'].upper()}] {a.get('code','')} : {a['message']}")
    
    # Check 02.5.3.2 in merged
    in_merged = merged02[merged02["Code"] == code_target]
    print(f"\n{code_target} dans merged_df: {'PRESENT' if not in_merged.empty else 'ABSENT'}")
    if not in_merged.empty:
        r = in_merged.iloc[0]
        print(f"  row_type={r['row_type']} | {col02}={r.get(col02)} | parent_code={r.get('parent_code','')}")
    
    # SALLE POLYVALENTE section total
    salle_merged = merged02[merged02["Désignation"].str.contains("POLYVALENTE|POLIVAL", case=False, na=False)]
    print(f"\nSALLE POLYVALENTE dans merged_df: {len(salle_merged)}")
    for _, r in salle_merged.iterrows():
        print(f"  Code={r['Code']} | type={r['row_type']} | {col02}={r.get(col02)} | Entete={repr(r['Entete'])}")
    
    # All 02.5.3.* in merged
    section_025_merged = merged02[merged02["Code"].str.startswith("02.5.3", na=False)]
    print(f"\nLignes 02.5.3.* dans merged_df:")
    for _, r in section_025_merged.iterrows():
        print(f"  Code={r['Code']:20s} | type={r['row_type']:15s} | {col02}={r.get(col02)}")
    
    print(f"\nTotaux finaux LOT 02:")
    for _, r in merged02[merged02["row_type"] == "total_line"].iterrows():
        print(f"  {r['Désignation'][:50]} : {r.get(col02)}")
        
except Exception as e:
    print(f"Erreur TCO LOT 02: {e}")
    import traceback; traceback.print_exc()
