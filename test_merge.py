import traceback
from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco

try:
    print('Loading TCO model...')
    tco_file = 'DPGF LOT 06 - ELECTRICITE.xlsx'
    tco_df, tco_meta = parse_tco(tco_file)
    print(f'TCO loaded with {len(tco_df)} rows')

    print('Loading company DPGF...')
    dpgf_file = 'DEVIS DPGF GUSTAVE EIFFEL BDX - Exemplaire Client.xlsx'
    dpgf_df, alerts = parse_dpgf(dpgf_file)
    print(f'DPGF loaded with {len(dpgf_df)} rows')

    print('Merging...')
    merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, 'EIFFEL BDX', tva_rate=20.0)
    
    # check if prices were updated
    company_cols = [c for c in merged_df.columns if 'EIFFEL BDX' in c]
    print(f'Merged cols: {company_cols}')
    if company_cols:
        prices_col = next(c for c in company_cols if '_Total' in c)
        total_sum = merged_df[prices_col].sum()
        print(f'Total company SUM: {total_sum}')
        
        # how many matched
        matched = merged_df[merged_df[prices_col] > 0]
        print(f'Matched rows with prices > 0: {len(matched)}')
        
except Exception as e:
    print('Error during testing:')
    traceback.print_exc()
