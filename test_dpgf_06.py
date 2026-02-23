import traceback
from core.parser_dpgf import parse_dpgf

try:
    df, alerts = parse_dpgf('DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx')
    print(df.head())
    print(f'Parsed {len(df)} rows')
except Exception as e:
    print('Error:')
    traceback.print_exc()
