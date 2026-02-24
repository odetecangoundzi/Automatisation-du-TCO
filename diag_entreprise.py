import pandas as pd
import os

filepath = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\entreprise\DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx"

try:
    # Read without header to see raw structure
    df_raw = pd.read_excel(filepath, header=None)
    print("Raw head (first 15 rows):")
    print(df_raw.head(15).to_string())
    
    print("\nRaw tail (last 15 rows):")
    print(df_raw.tail(15).to_string())
    
    # Check for keywords in Column B (Désignation usually)
    print("\nRows matching 'TOTAL', 'TVA', 'TTC' in column 1 (B):")
    mask = df_raw[1].astype(str).str.contains("TOTAL|TVA|TTC", case=False, na=False)
    print(df_raw[mask].to_string())

    # Check column 5 (F) where Px Tot HT is usually located
    print("\nValues in column 5 (F) for those rows:")
    print(df_raw[mask][5].to_string())

except Exception as e:
    print(f"Error reading {filepath}: {e}")
