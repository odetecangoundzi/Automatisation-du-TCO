"""
utils.py — Fonctions partagées entre les parsers TCO et DPGF.

Centralise :
  - _find_header_row  : détecte la ligne d'en-tête Code|Désignation
  - _classify_row     : classifie chaque ligne selon la colonne Entete (col M)
"""

from __future__ import annotations
import pandas as pd


def find_header_row(df: pd.DataFrame, max_search: int = 20) -> int:
    """
    Parcourt les lignes d'un DataFrame pour trouver celle contenant
    'Code' en col 0 et 'Désignation' en col 1.

    Returns:
        int : index de la ligne (0-indexed)
    Raises:
        ValueError : si l'en-tête n'est pas trouvée dans les max_search premières lignes
    """
    for row_idx in range(min(len(df), max_search)):
        row = df.iloc[row_idx]
        if len(row) < 2:
            continue
            
        a_val = row.iloc[0]
        b_val = row.iloc[1]
        
        if (
            pd.notna(a_val) and str(a_val).strip().lower() == "code"
            and pd.notna(b_val) and "signation" in str(b_val).strip().lower()
        ):
            return row_idx
            
    raise ValueError(
        "Impossible de trouver la ligne d'en-tête (Code | Désignation) "
        f"dans les {max_search} premières lignes."
    )


def classify_row(code_str: str, desig_str: str, entete_str: str, has_price: bool = False) -> str:
    """
    Classifie une ligne selon les métadonnées de la colonne Entete (col M).
    Si l'entête est absente ou non standard, utilise le paramètre `has_price` en fallback
    pour identifier les articles.

    Types retournés :
      - section_header : section principale (Bd_xxx_Bord)
      - recap          : totalisation par section (Bord_xxx_Recap, Code vide)
      - recap_summary  : table récap en fin de fichier (RecapBord_xxx)
      - sub_section    : sous-section (Ouv_xxx_Niv1 / Niv2)
      - article        : ligne de détail avec prix (Ouv_xxx_Art)
      - total_line     : ligne Montant HT / TVA / TTC (LignesTot_xxx)
      - total_text     : ligne dont le code commence par 'Total'
      - empty          : ligne sans code ni désignation
      - other          : tout le reste non reconnu

    Args:
        code_str   (str) : valeur de la colonne Code (déjà strip())
        desig_str  (str) : valeur de la colonne Désignation
        entete_str (str) : valeur de la colonne Entete (col M)
        has_price (bool) : True si la ligne a une quantité et un prix unitaire

    Returns:
        str : type de ligne
    """
    ent = entete_str

    if "RecapBord" in ent:
        return "recap_summary"
    if "LignesTot" in ent:
        return "total_line"
    
    # Fallback si l'entête est manquante (cas fréquent sur lignes de total)
    d_low = desig_str.lower()
    if "montant ht" in d_low or ("tva" in d_low and "ht" not in d_low) or "montant ttc" in d_low:
        return "total_line"
    if "Bord" in ent and "Recap" in ent:
        return "recap"
    if ent.startswith("Bd_") and "Bord" in ent:
        return "section_header"
    if "_Niv1" in ent or "_Niv2" in ent:
        return "sub_section"
    if "_Art" in ent:
        return "article"
    if code_str.lower().startswith("total"):
        return "total_text"
    if not code_str and not desig_str:
        return "empty"
    
    # Fallback pour les DPGF sans colonne Entete remplie
    if has_price and code_str:
        return "article"
        
    return "other"
