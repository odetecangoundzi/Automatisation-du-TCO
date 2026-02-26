"""
utils.py — Fonctions partagées entre les parsers TCO et DPGF.

Centralise :
  - find_header_row : détecte la ligne d'en-tête Code|Désignation
  - find_column_index : mappe les colonnes par mots-clés
  - classify_row : classifie chaque ligne selon la colonne Entete (col M)
"""

from __future__ import annotations
import pandas as pd


def find_header_row(df: pd.DataFrame, max_search: int = 40) -> int:
    """
    Parcourt les lignes d'un DataFrame pour trouver celle contenant
    'Code' et 'Désignation' (ou leurs équivalents hétérogènes).

    Synonymes acceptés pour la colonne Code :
      "code", "n°", "n°.", "num", "indice", "ref", "no"
    Synonymes acceptés pour la colonne Désignation :
      sous-chaîne "signation", "libellé", "libelle"

    Fallback (DPGFs sans colonne Code, ex : ERTIE&FILS) :
      "Désignation" présente + au moins une colonne prix/unité reconnue.
    """
    _CODE_SYNONYMS = frozenset(
        {"code", "n°", "n°.", "num", "indice", "ref", "no"}
    )
    _PRICE_MARKERS = ("p.u", "px u", "prix u", "montant", "total h", "h.t.")

    for row_idx in range(min(len(df), max_search)):
        row = [str(val).strip().lower() for val in df.iloc[row_idx]]
        if len(row) < 2:
            continue

        has_code = any(val in _CODE_SYNONYMS for val in row[:5])
        has_desig = any(
            "signation" in val
            or "libellé" in val
            or "libelle" in val
            for val in row[:6]
        )

        if has_code and has_desig:
            return row_idx

    # Fallback : DPGFs sans colonne "Code" explicite (ex : ERTIE&FILS).
    # Accepté si "Désignation" + au moins une colonne prix/unité reconnue.
    for row_idx in range(min(len(df), max_search)):
        row = [str(val).strip().lower() for val in df.iloc[row_idx]]
        has_desig = any(
            "signation" in val
            or "libellé" in val
            or "libelle" in val
            for val in row[:6]
        )
        has_price_header = any(
            any(marker in val for marker in _PRICE_MARKERS)
            for val in row
        )
        if has_desig and has_price_header:
            return row_idx

    raise ValueError(
        "Impossible de trouver la ligne d'en-tête (Code | Désignation) "
        f"dans les {max_search} premières lignes."
    )


def find_column_index(
    df: pd.DataFrame, keywords: list[str], default_idx: int
) -> int:
    """
    Cherche l'index d'une colonne par correspondance de mots-clés.
    Si non trouvé, retourne l'index par défaut.

    Règle de matching :
      - mot-clé de 1 caractère → correspondance exacte (avec/sans point)
        ex: "u" matche "u." mais PAS "qu. ent."
      - mot-clé de 2+ caractères → correspondance par sous-chaîne
    """
    cols = [str(c).strip().lower() for c in df.columns]
    for i, col in enumerate(cols):
        col_base = col.rstrip(". ")     # "u." → "u", "qu. ent." → "qu. ent"
        for kw in keywords:
            kw_l = kw.lower()
            kw_base = kw_l.rstrip(". ")  # "qu." → "qu", "u" → "u"
            if kw_l == col or kw_base == col_base:
                return i
            if len(kw_l) > 1 and kw_l in col:
                return i
    return default_idx


def classify_row(
    code_str: str,
    desig_str: str,
    entete_str: str,
    has_price: bool = False,
) -> str:
    """
    Classifie une ligne selon les métadonnées de la colonne Entete (col M).
    Si l'entête est absente ou non standard, utilise des heuristiques
    sur le code et la désignation en fallback.

    Types retournés :
      - section_header : section principale (Bd_xxx_Bord ou code '01.1')
      - recap          : totalisation par section (Bord_xxx_Recap)
      - recap_summary  : table récap en fin de fichier (RecapBord_xxx)
      - sub_section    : sous-section (Ouv_xxx_Niv1/Niv2 ou '01.1.1')
      - article        : ligne de détail avec prix (Ouv_xxx_Art ou has_price)
      - total_line     : Montant HT / TVA / TTC (LignesTot_xxx)
      - total_text     : ligne dont le code commence par 'Total'
      - empty          : ligne sans code ni désignation
      - other          : tout le reste non reconnu
    """
    ent = entete_str
    code = code_str.lower()
    desig = desig_str.lower()

    # 1. Détection via Entete (priorité haute)
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

    # 2. Détection via Désignation (fallback totaux)
    if (
        "montant ht" in desig
        or ("tva" in desig and "ht" not in desig)
        or "montant ttc" in desig
    ):
        return "total_line"
    if "total" in desig and ("section" in desig or "lot" in desig):
        return "recap"

    # 3. Détection via structure du Code
    if code.startswith("total"):
        return "total_text"

    # Heuristique : nombre de points dans le code
    # ex: 01.1 = section, 01.1.1 = sub_section, 01.1.1.1 = article
    parts = [p for p in code.split(".") if p.strip()]
    if parts:
        if len(parts) == 2:
            return "section_header"
        if len(parts) == 3:
            return "sub_section"
        if len(parts) >= 4:
            return "article"

    # 4. Fallback vide
    if not code_str and not desig_str:
        return "empty"

    # 5. Fallback avec prix (code peut être absent sur certains formats DPGF)
    if has_price:
        return "article"

    return "other"
