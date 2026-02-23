from core.parser_dpgf import parse_dpgf
import pandas as pd

dpgf_file = 'DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx'
dpgf_df, alerts = parse_dpgf(dpgf_file)

with_prices = dpgf_df[dpgf_df['Px_Tot_HT'].apply(lambda x: isinstance(x, (int, float)) and x > 0)]

print("Top 20 rows with prices (from raw excel):")
for idx, row in with_prices.head(20).iterrows():
    print(f"Row {row['original_row']}: Code='{row['Code']}' | Desig='{row['Désignation'][:40]}...' | Qu='{row['Qu.']}' | PU='{row['Px_U_HT']}' | Tot='{row['Px_Tot_HT']}'")

