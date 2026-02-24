import sys
import os
from decimal import Decimal
import pandas as pd

# Add current dir to path to import core
sys.path.append(os.getcwd())

from core.utils import classify_row
from core.parser_dpgf import _clean_numeric
from core.parser_tco import _extract_project_info

def test_p1_classification():
    print("Testing P1: Classification Fallbacks")
    # Test section header by code structure
    assert classify_row("01.1", "Some Section", "") == "section_header"
    # Test sub section by code structure
    assert classify_row("01.1.1", "Some Sub", "") == "sub_section"
    # Test article by code structure
    assert classify_row("01.1.1.1", "Some Art", "") == "article"
    # Test recap by designation
    assert classify_row("", "Total de la section", "") == "recap"
    print("  [OK] P1 Classification fallbacks verified.")

def test_p2_regex():
    print("Testing P2: Regex Numeric Precision")
    # Test value with lot prefix
    val1 = "Lot 1 : 1250.50 €"
    num, comm = _clean_numeric(val1)
    assert num == Decimal("1250.50")
    
    # Test value with index at start
    val2 = "1. 4500"
    num, comm = _clean_numeric(val2)
    assert num == Decimal("4500")
    
    # Test complex text
    val3 = "Option A (Ref 123) : 99.99"
    num, comm = _clean_numeric(val3)
    assert num == Decimal("99.99")
    print("  [OK] P2 Regex numeric precision verified (Last number extracted).")

def test_p4_project_info():
    print("Testing P4: Project Info Keyword Extraction")
    df = pd.DataFrame([
        ["PROJET: SUPER CHANTIER", None],
        [None, None],
        ["LOT n°04", "PLATRERIE"]
    ])
    info = _extract_project_info(df, 3)
    assert info.get("projet") == "SUPER CHANTIER"
    assert info.get("lot") == "PLATRERIE"
    print("  [OK] P4 Project info keyword extraction verified.")

if __name__ == "__main__":
    try:
        test_p1_classification()
        test_p2_regex()
        test_p4_project_info()
        print("\nALL CORE FIXES VERIFIED SUCCESSFULLY!")
    except Exception as e:
        print(f"\n[FAILURE] Verification failed: {e}")
        sys.exit(1)
