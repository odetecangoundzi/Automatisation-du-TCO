from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco

tco_file = 'DPGF LOT 06 - ELECTRICITE.xlsx'
tco_df, tco_meta = parse_tco(tco_file)

dpgf_file = 'DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx'
dpgf_df, alerts = parse_dpgf(dpgf_file)

merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, 'EIFFEL BDX', tva_rate=20.0)

matches = merged_df[merged_df['EIFFEL BDX_Px_Tot_HT'] > 0]
print(f"Number of rows with Px_Tot_HT > 0: {len(matches)}")
if len(matches) == 0:
    print("Zero matches! Let's check the codes in TCO and DPGF:")
    print("TCO Codes (top 10):", tco_df[tco_df['row_type'] == 'article']['Code'].head(10).tolist())
    print("DPGF Codes (top 10):", dpgf_df[dpgf_df['row_type'] == 'article']['Code'].head(10).tolist())
