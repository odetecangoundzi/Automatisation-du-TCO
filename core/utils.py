"""
utils.py — Fonctions partagées entre les parsers TCO et DPGF.

Centralise :
  - _find_header_row  : détecte la ligne d'en-tête Code|Désignation
  - _classify_row     : classifie chaque ligne selon la colonne Entete (col M)
"""


def find_header_row(ws, max_search=20):
    """
    Parcourt les lignes du worksheet pour trouver celle contenant
    'Code' en col A et 'Désignation' en col B.

    Returns:
        int : numéro de ligne (1-indexed)
    Raises:
        ValueError : si l'en-tête n'est pas trouvée dans les max_search premières lignes
    """
    for row_idx in range(1, min(ws.max_row + 1, max_search + 1)):
        a_val = ws.cell(row=row_idx, column=1).value
        b_val = ws.cell(row=row_idx, column=2).value
        if (
            a_val and str(a_val).strip().lower() == "code"
            and b_val and "signation" in str(b_val).strip().lower()
        ):
            return row_idx
    raise ValueError(
        "Impossible de trouver la ligne d'en-tête (Code | Désignation) "
        f"dans les {max_search} premières lignes."
    )


def classify_row(code_str, desig_str, entete_str):
    """
    Classifie une ligne selon les métadonnées de la colonne Entete (col M).

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

    Returns:
        str : type de ligne
    """
    ent = entete_str

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
    if code_str.lower().startswith("total"):
        return "total_text"
    if not code_str and not desig_str:
        return "empty"
    return "other"
