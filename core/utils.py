"""
utils.py — Fonctions partagées entre les parsers TCO et DPGF.

Centralise :
  - find_header_row  : détecte la ligne d'en-tête Code|Désignation
  - find_column_index : recherche de colonne par mots-clés
  - classify_row     : classifie chaque ligne selon la colonne Entete (col M)
"""

from __future__ import annotations

import pandas as pd

# Sentinel retourné par find_column_index quand la colonne n'est pas trouvée
# et qu'aucun default_idx n'est fourni.
COL_NOT_FOUND = -1


def find_header_row(df: pd.DataFrame, max_search: int = 40) -> int:
    """
    Parcourt les lignes d'un DataFrame pour trouver celle contenant
    'Code' et 'Désignation'.
    """
    for row_idx in range(min(len(df), max_search)):
        row = [str(val).strip().lower() for val in df.iloc[row_idx]]
        if len(row) < 2:
            continue

        has_code = any("code" == val for val in row[:3])
        has_desig = any("signation" in val or "libellé" in val for val in row[:4])

        if has_code and has_desig:
            return row_idx

    raise ValueError(
        "Impossible de trouver la ligne d'en-tête (Code | Désignation) "
        f"dans les {max_search} premières lignes."
    )


def find_column_index(
    df: pd.DataFrame,
    keywords: list[str],
    default_idx: int | None = None,
) -> int:
    """
    Cherche l'index d'une colonne par correspondance de mots-clés dans les noms de colonnes.

    Retourne :
        - L'index de la première colonne qui correspond à un mot-clé.
        - default_idx si non trouvé et default_idx est fourni (backward compat).
        - COL_NOT_FOUND (-1) si non trouvé et default_idx est None.

    Règle de matching :
      - mot-clé de 1 caractère → correspondance exacte uniquement (évite que "u" matche "quantité")
      - mot-clé de 2+ caractères → correspondance par sous-chaîne (comportement historique)
    """
    cols = [str(c).strip().lower() for c in df.columns]
    for i, col in enumerate(cols):
        for kw in keywords:
            kw_l = kw.lower()
            if kw_l == col or (len(kw_l) > 1 and kw_l in col):
                return i
    if default_idx is None:
        return COL_NOT_FOUND
    return default_idx


def classify_row(code_str: str, desig_str: str, entete_str: str, has_price: bool = False) -> str:
    """
    Classifie une ligne selon les métadonnées de la colonne Entete (col M).
    Si l'entête est absente ou non standard, utilise des heuristiques sur le code
    et la désignation en fallback.

    Types retournés :
      - section_header : section principale (Bd_xxx_Bord ou code court '01.1')
      - recap          : totalisation par section (Bord_xxx_Recap, ou 'Total section')
      - recap_summary  : table récap en fin de fichier (RecapBord_xxx)
      - sub_section    : sous-section (Ouv_xxx_Niv1 / Niv2 ou code '01.1.1')
      - article        : ligne de détail avec prix (Ouv_xxx_Art ou has_price)
      - total_line     : ligne Montant HT / TVA / TTC (LignesTot_xxx)
      - total_text     : ligne dont le code commence par 'Total'
      - empty          : ligne sans code ni désignation
      - other          : tout le reste non reconnu
    """
    ent = entete_str
    code = code_str.lower()
    desig = desig_str.lower()

    # 1. Détection via Entete (Priorité haute)
    if "RecapBord" in ent:
        return "recap_summary"
    if "LignesTot" in ent:
        return "total_line"
    if "Bord" in ent and "Recap" in ent:
        return "recap"
    if ent.startswith("Bd_") and "Bord" in ent:
        return "section_header"
    if "_Niv1" in ent or "_Niv2" in ent:
        return "sub_section"
    if "_Art" in ent:
        return "article"

    # 2. Détection via Désignation (Fallback Totaux)
    if "montant ht" in desig or ("tva" in desig and "ht" not in desig) or "montant ttc" in desig:
        return "total_line"
    if "total" in desig and ("section" in desig or "lot" in desig):
        return "recap"

    # 3. Priorité prix : une ligne avec Qu et PU renseignés est forcément un article
    #    (les section_headers n'ont jamais de prix directs — valeurs calculées)
    if has_price and code_str:
        return "article"

    # 4. Détection via Structure du Code (fallback sans prix)
    if code.startswith("total"):
        return "total_text"

    parts = [p for p in code.split(".") if p.strip()]
    if parts:
        if len(parts) <= 2:  # "01" ou "01.1" → section principale
            return "section_header"
        if len(parts) == 3:
            return "sub_section"
        if len(parts) >= 4:
            return "article"

    # 5. Fallback Vide
    if not code_str and not desig_str:
        return "empty"

    return "other"
