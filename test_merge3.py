from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
import pandas as pd

dpgf_file = 'DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx'
dpgf_df, alerts = parse_dpgf(dpgf_file)

# Show all types of rows
print(dpgf_df['row_type'].value_counts())

# print out empty codes
no_code = dpgf_df[dpgf_df['Code'].astype(str).str.strip() == '']
print(f"Num rows with no code: {len(no_code)}")

# Let's inspect rows with actual prices to see how they are marked:
with_prices = dpgf_df[dpgf_df['Px_Tot_HT'].apply(lambda x: isinstance(x, (int, float)) and x > 0)]
print(f"Num rows with positive prices: {len(with_prices)}")
print("Categories of rows with positive prices:")
print(with_prices['row_type'].value_counts())

print("\nSample of rows with positive prices:")
print(with_prices[['Code', 'Désignation', 'row_type', 'Px_Tot_HT']].head(10))

# What is the entete for these rows?
print("\nSample Entetes for rows with prices:")
print(with_prices['Entete'].head(10).tolist())
