import os
import pandas as pd
import openpyxl
from core.parser_tco import parse_tco
from core.parser_dpgf import parse_dpgf
from core.merger import merge_company_into_tco

def create_mock_tco(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    # Minimal TCO structure - Header row must match find_header_row logic (contains 'code' and 'designation')
    ws.append(["CODE", "DESIGNATION", "Quantité", "Unité", "Px U HT", "Px Tot HT", "Entête"])
    # Codes with 4 parts are identified as articles by classify_row
    ws.append(["01.01.01", "SECTION 1", None, None, None, None, "Bd_Bord"])
    ws.append(["01.01.01.01", "Article 1", 1, "m2", 100, 100, "Ouv_Art"])
    ws.append(["01.01.01.02", "Article 2", 2, "u", 50, 100, "Ouv_Art"])
    wb.save(fname)

def create_mock_dpgf(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    # Minimal DPGF structure
    ws.append(["CODE", "DESIGNATION", "Qu.", "U", "Px U", "Px Tot"])
    ws.append(["01.01.01.01", "Article 1", 1, "m2", 110, 110])
    ws.append(["01.01.01.02", "Article 2", 2, "u", 60, 120])
    wb.save(fname)

def test_basic_merge():
    tco_file = 'mock_tco.xlsx'
    dpgf_file = 'mock_dpgf.xlsx'
    
    try:
        print('Creating mock files...')
        create_mock_tco(tco_file)
        create_mock_dpgf(dpgf_file)
        
        print('Loading TCO model...')
        tco_df, tco_meta = parse_tco(tco_file)
        print(f"TCO loaded: {len(tco_df)} rows, "
              f"{len(tco_df[tco_df['row_type']=='article'])} articles")
        
        print('Loading company DPGF...')
        dpgf_df, alerts = parse_dpgf(dpgf_file)
        
        print('Merging...')
        merged_df, merge_alerts = merge_company_into_tco(tco_df, dpgf_df, 'COMPANY_A', tva_rate=20.0)
        
        # check if prices were updated
        company_cols = [c for c in merged_df.columns if 'COMPANY_A' in c]
        assert len(company_cols) > 0, "No company columns found after merge"
        
        # The merger uses _Px_Tot_HT as suffix
        prices_col = next(c for c in company_cols if '_Px_Tot_HT' in c)
        total_sum = merged_df[prices_col].sum()
        print(f'Total company SUM: {total_sum}')
        assert total_sum == 230, f"Expected 230, got {total_sum}"
        
        # how many matched
        matched = merged_df[merged_df[prices_col] > 0]
        assert len(matched) == 2, f"Expected 2 matches, got {len(matched)}"

    finally:
        for f in [tco_file, dpgf_file]:
            if os.path.exists(f):
                os.remove(f)

if __name__ == "__main__":
    import pytest
    pytest.main([__file__])
