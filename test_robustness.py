import os
import traceback
import pandas as pd
from core.parser_dpgf import parse_dpgf
from core.parser_tco import parse_tco
import openpyxl

print("=== DEBUT DES TESTS DE ROBUSTESSE EXCEL ===")

def test_file(filename, description, setup_fn):
    print(f"\n--- Test: {description} ---")
    try:
        setup_fn(filename)
        print("  Fichier cree.")
        try:
            df, alerts = parse_dpgf(filename)
            print(f"  [OK] Parse {len(df)} lignes, {len(alerts)} alertes.")
        except Exception as e:
            print(f"  [ERROR] {type(e).__name__} - {e}")
    finally:
        try:
            if os.path.exists(filename):
                os.remove(filename)
        except OSError as e:
            print(f"  [WARN] Impossible de supprimer {filename}: {e}")

# 1. Fichier complètement vide
def create_empty(fname):
    wb = openpyxl.Workbook()
    wb.save(fname)

test_file("test_empty.xlsx", "Fichier Excel vide (0 ligne, 0 colonne)", create_empty)

# 2. Fichier avec des données mais pas d'en-tête (Code | Désignation)
def create_no_header(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Coucou", "Ceci", "N'est", "Pas", "Un", "DPGF"])
    ws.append(["123", "456", "789"])
    wb.save(fname)

test_file("test_no_header.xlsx", "Fichier sans en-tete standard", create_no_header)

# 3. Fichier avec en-tête mais sans lignes de données
def create_only_header(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Code", "Designation", "Qu.", "U", "Px U", "Px Tot"])
    wb.save(fname)

test_file("test_only_header.xlsx", "Fichier avec entete mais sans donnees", create_only_header)

# 4. Fichier avec en-tête mais colonnes décalées (ex: 3 colonnes vides au début)
def create_shifted_header(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["", "", "", "Code", "Designation", "Qu.", "U", "Px U", "Px Tot"])
    ws.append(["", "", "", "01.1", "Truc", 1, "m2", 10, 10])
    wb.save(fname)

test_file("test_shifted.xlsx", "Colonnes decalees horizontalement", create_shifted_header)

# 5. Fichier avec lignes de données tronquées (pas de colonne prix)
def create_truncated(fname):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Code", "Designation", "Qu.", "U", "Px U", "Px Tot"]) 
    ws.append(["01.1", "Article sans prix"]) 
    wb.save(fname)

test_file("test_truncated.xlsx", "Fichier sans les colonnes de prix/quantite", create_truncated)

print("\n=== FIN DES TESTS ===")
