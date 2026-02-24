import pandas as pd

lot01_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 01 - DESAMIANTAGE - CURAGE - GROS OEUVRE.xlsx"
lot06_path = r"s:\TRANSFORMATION DIGITALE\Automatisation du TCO\TCO_APP\DPGF LOT 06 - ELECTRICITE.xlsx"

try:
    df01 = pd.read_excel(lot01_path, header=None)
    print("LOT 01 Codes (Column A):")
    print(df01[0].dropna().head(10).to_string())
    
    df06 = pd.read_excel(lot06_path, header=None)
    print("\nLOT 06 Codes (Column A):")
    print(df06[0].dropna().head(10).to_string())

except Exception as e:
    print(f"Error: {e}")
