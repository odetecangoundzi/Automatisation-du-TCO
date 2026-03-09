"""
utils.py — Fonctions partagées entre les parsers TCO et DPGF.

Centralise :
  - find_header_row   : détecte la ligne d'en-tête Code|Désignation
  - find_column_index : recherche de colonne par mots-clés
  - classify_row      : classifie chaque ligne selon la colonne Entete (col M)
  - open_excel_file   : ouvre un fichier Excel, détecte la feuille et l'en-tête
"""

from __future__ import annotations

import re
import os

import pandas as pd

# Regex centralisées pour la détection des options et variantes
# Déclencheurs sur la désignation (singulier/pluriel, majuscule/minuscule)
_RE_OPTION_DESIG = re.compile(
    r"\b(options?|variantes?|variante\s+libre|variante\s+imposee|pse|supplément|supplémentaire|suppl\.?)\b",
    re.I,
)
# Codes de type OPT, OPT1, OPT2 ou commençant par VAR
_RE_OPTION_CODE = re.compile(r"^(OPT\d*|VAR.*|.*\.VAR|.*\.OPT)$", re.I)

# Sentinel retourné par find_column_index quand la colonne n'est pas trouvée
# et qu'aucun default_idx n'est fourni.
COL_NOT_FOUND = -1


def open_excel_file(
    filepath: str,
) -> tuple[pd.ExcelFile, str, pd.DataFrame, int, dict]:
    """
    Ouvre un fichier Excel, détecte la bonne feuille et la ligne d'en-tête.

    Factorise la logique commune à parser_tco et parser_dpgf :
      1. Détermine le moteur (openpyxl / xlrd / pyxlsb) selon l'extension
      2. Ouvre le fichier une seule fois (pd.ExcelFile)
      3. Sonde chaque feuille pour trouver celle avec un en-tête valide
      4. Réutilise le DataFrame probe (évite une 2ᵉ lecture disque)
      5. Fait une unique lecture finale avec skiprows pour df_data

    Returns:
        xl_file        : pd.ExcelFile ouvert (à fermer par l'appelant)
        sheet_name     : nom de la feuille retenue
        df_raw         : DataFrame header=None de la feuille (réutilisé depuis le probe)
        header_row_idx : index (0-based) de la ligne d'en-tête dans df_raw
        engine_kwargs  : kwargs moteur (ex: {"data_only": True} pour openpyxl)

    Raises:
        ValueError si aucune feuille valide n'est trouvée.
    """
    ext = os.path.splitext(filepath)[1].lower()
    engine: str | None = None
    if ext == ".xls":
        engine = "xlrd"
    elif ext == ".xlsb":
        engine = "pyxlsb"
    elif ext in (".xlsx", ".xlsm"):
        engine = "openpyxl"

    # data_only=True : lit les valeurs calculées des formules (évite "=C5*E5" dans les cellules)
    engine_kwargs: dict = {"data_only": True} if engine == "openpyxl" else {}

    # Ouverture unique — les feuilles sont listées sans tout lire
    # Fallback xlrd si openpyxl échoue (ex : fichier .xls renommé en .xlsx,
    # ou fichier généré par un logiciel non conforme OOXML).
    try:
        xl_file = pd.ExcelFile(filepath, engine=engine, engine_kwargs=engine_kwargs)
    except Exception as _first_err:
        if engine == "openpyxl":
            try:
                xl_file = pd.ExcelFile(filepath, engine="xlrd")
                engine_kwargs = {}  # xlrd n'accepte pas data_only
            except Exception:
                raise _first_err from None
        else:
            raise
    all_sheets = xl_file.sheet_names

    sheet_name = all_sheets[0]
    df_raw: pd.DataFrame | None = None

    # Sondage : trouver la feuille avec un en-tête Code|Désignation valide
    for sn in all_sheets:
        df_probe = xl_file.parse(sn, header=None)
        try:
            find_header_row(df_probe)
            sheet_name = sn
            df_raw = df_probe  # Réutilisation : évite une 2ᵉ lecture
            break
        except ValueError:
            continue

    if df_raw is None:
        # Aucune feuille avec en-tête valide — on garde la première et on laisse
        # find_header_row lever l'erreur avec un message clair
        df_raw = xl_file.parse(all_sheets[0], header=None)

    header_row_idx = find_header_row(df_raw)
    return xl_file, sheet_name, df_raw, header_row_idx, engine_kwargs


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
    _CODE_SYNONYMS = frozenset({"code", "n°", "n°.", "num", "indice", "ref", "no"})
    _PRICE_MARKERS = ("p.u", "px u", "prix u", "montant", "total h", "h.t.")

    for row_idx in range(min(len(df), max_search)):
        row = [str(val).strip().lower() for val in df.iloc[row_idx]]
        if len(row) < 2:
            continue

        has_code = any(val in _CODE_SYNONYMS for val in row[:5])
        has_desig = any(
            "signation" in val or "libellé" in val or "libelle" in val for val in row[:6]
        )

        if has_code and has_desig:
            return row_idx

    # Fallback : DPGFs sans colonne "Code" explicite (ex : ERTIE&FILS).
    # Accepté si "Désignation" + au moins une colonne prix/unité reconnue.
    for row_idx in range(min(len(df), max_search)):
        row = [str(val).strip().lower() for val in df.iloc[row_idx]]
        has_desig = any(
            "signation" in val or "libellé" in val or "libelle" in val for val in row[:6]
        )
        has_price_header = any(any(marker in val for marker in _PRICE_MARKERS) for val in row)
        if has_desig and has_price_header:
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
      - mot-clé de 1 caractère → correspondance exacte (avec/sans point)
        ex: "u" matche "u." mais PAS "qu. ent."
      - mot-clé de 2+ caractères → correspondance par sous-chaîne
    """
    cols = [str(c).strip().lower() for c in df.columns]
    for i, col in enumerate(cols):
        col_base = col.rstrip(". ")  # "u." → "u", "qu. ent." → "qu. ent"
        for kw in keywords:
            kw_l = kw.lower()
            kw_base = kw_l.rstrip(". ")  # "qu." → "qu", "u" → "u"
            if kw_l == col or kw_base == col_base:
                return i
            if len(kw_l) > 1 and kw_l in col:
                return i
    if default_idx is None:
        return COL_NOT_FOUND
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

    # 3. Priorité prix : une ligne avec Qu et PU renseignés est forcément un article,
    #    quelle que soit la profondeur du code (2, 3 ou 4 segments).
    #    Les section_headers n'ont jamais de prix directs (valeurs calculées).
    if has_price and code_str:
        return "article"

    # 4. Détection via Structure du Code (fallback sans prix, sans Entete).
    #    Règle permissive : on ne présume pas qu'un LOT a toujours N niveaux fixes.
    #    On utilise le nombre de segments uniquement quand has_price est False,
    #    donc sans risque de confondre un article à 2 segments avec un section_header.
    if code.startswith("total"):
        return "total_text"

    parts = [p for p in code.split(".") if p.strip()]
    if parts:
        n = len(parts)
        if n == 1:
            # Code court sans point → section principale (ex : "01", "A", "I")
            return "section_header"
        if n == 2:
            # Ex : "01.1" — sans prix, traité comme section (chapitre intermédiaire)
            # Cas couverts par has_price=True déjà retournés "article" au-dessus.
            return "section_header"
        if n == 3:
            # Ex : "01.1.2" — sous-section ou article sans prix explicite
            return "sub_section"
        # n >= 4 → article profondément imbriqué
        return "article"

    # 5. Fallback Vide
    if not code_str and not desig_str:
        return "empty"

    # 6. Fallback avec prix sans code (ex : ERTIE&FILS — pas de colonne Code)
    if has_price:
        return "article"

    return "other"


def is_option_row(code_str: str, desig_str: str) -> bool:
    """
    Détermine si une ligne correspond à une option ou une variante
    en basant la détection sur le code et la désignation.
    """
    code = str(code_str or "").strip()
    desig = str(desig_str or "").strip().lower()

    if not code and not desig:
        return False

    # 1. Déclencheur sur la désignation
    if _RE_OPTION_DESIG.search(desig):
        return True

    # 2. Déclencheur sur le code
    if code and _RE_OPTION_CODE.match(code):
        return True

    return False
