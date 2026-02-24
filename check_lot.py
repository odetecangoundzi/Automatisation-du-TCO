import pandas as pd
import sys

filepath = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx"

try:
    df_raw = pd.read_excel(filepath, header=None)
    print("Project info rows (first 10):")
    print(df_raw.head(10).to_string())
    
    # Look for "LOT" keyword
    mask = df_raw.astype(str).apply(lambda x: x.str.contains("LOT", case=False)).any(axis=1)
    print("\nRows containing 'LOT':")
    print(df_raw[mask].to_string())
    
    # Look at some codes
    print("\nFirst 10 codes (Column A):")
    print(df_raw[0].dropna().head(20).to_string())

except Exception as e:
    print(f"Error: {e}")
