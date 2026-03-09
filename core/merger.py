"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Calcule les sous-totaux en se basant sur la hiérarchie Code (01.X → 01.X.Y...).
Les lignes "recap" (Code vide, Entete Bord_xxx_Recap) reçoivent le total
de leur section parente.
"""

import re
from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

import pandas as pd

from config import TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)

_RE_MONTANT_HT = re.compile(r"montant\s+ht")
_RE_TVA_ONLY = re.compile(r"\btva\b")
_RE_HT_ONLY = re.compile(r"\bht\b")
_RE_MONTANT_TTC = re.compile(r"montant\s+ttc|(?<!\w)ttc(?!\w)")
_RE_JUNK_TOTAL = re.compile(
    r"\b(total|sous-total|montant|somme|global|net\s+a\s+payer|tva|ttc)\b", re.I
)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _normalize_code(code: object) -> str:
    """Normalise un code pour la comparaison.

    Supprime les zéros de tête sur chaque segment du code hiérarchique.
    Traite les floats retournés par Excel pour des cellules numériques.

    Exemples :
        "01.10"    → "1.10"   (zéros de tête, segment 1)
        "01.1"     → "1.1"
        "01"       → "1"
        "03.5.2"   → "3.5.2"  (multi-niveaux → normalisation par segment)
        float 1.0  → "1"      (Excel lit "01" comme 1.0)
        float 1.1  → "1.1"    (Excel lit "01.1" comme 1.1)
    """
    if code is None:
        return ""
    if isinstance(code, float):
        if pd.isna(code):
            return ""
        # Float issu d'Excel : normaliser sans trailing zeros artificiels
        s = str(int(code)) if code == int(code) else str(code)
    else:
        s = str(code).strip()

    if not s or s.lower() in ("nan", "none"):
        return ""

    # Normalisation par segment : strip leading zeros, préserve trailing zeros
    # "01.10" → ["01","10"] → ["1","10"] → "1.10"  (≠ "1.1" pour "01.1")
    # "03.5.2" → ["03","5","2"] → ["3","5","2"] → "3.5.2"
    parts = s.split(".")
    return ".".join(p.lstrip("0") or "0" for p in parts)


def _detect_malformed_code(raw_code: object) -> tuple[bool, str]:
    """Détecte un code DPGF mal formé et tente une correction automatique.

    Un code est considéré malformé s'il contient :
    - Des virgules à la place des points  ("2.1,1,3")
    - Des espaces intempestifs            ("2. 1.3")
    - Des points doublés ou finaux        ("2..1", "2.1.")
    - Un suffixe alpha (variante)         ("2.4.1.5b", "2.4.1.5-bis")
    - Des caractères non alphanumériques  ("2.1.3/")

    Returns:
        (is_malformed, corrected_normalized_code)
        - is_malformed = True si le code d'origine est malformé
        - corrected_normalized_code = code corrigé et normalisé, ou "" si non corrigeable
    """
    import re as _re

    if raw_code is None:
        return False, ""
    s = str(raw_code).strip()
    if not s or s.lower() in ("nan", "none"):
        return False, ""

    # Pas de problème sur un code entier/float standard
    if isinstance(raw_code, float):
        return False, _normalize_code(raw_code)

    # Tentative de correction : remplacer virgules et espaces par des points
    corrected = s.replace(",", ".").replace(" ", "")

    # Supprimer les points doublés et finaux
    corrected = _re.sub(r"\.{2,}", ".", corrected).strip(".")

    # Vérifier la présence de caractères invalides dans le code corrigé
    # Un code valide est composé uniquement de chiffres et de points
    has_invalid_chars = bool(_re.search(r"[^\d.]", corrected))

    is_malformed = (s != corrected) or has_invalid_chars

    if has_invalid_chars:
        # Codes dupliqués générés par le merger (ex: 2.6.5.4_DUP02 → 2.6.5.4)
        _m_dup = _re.match(r"^(\d+(?:\.\d+)*)_DUP\d+$", s)
        if _m_dup:
            return True, _normalize_code(_m_dup.group(1))

        # Tentative de correction : suffixe alpha (variante technique)
        # Exemples : "2.4.1.5b" → "2.4.1.5" | "2.9.1.1-bis" → "2.9.1.1"
        _RE_VARIANT = _re.compile(r"^(\d+(?:\.\d+)*)-?([a-zA-Z]+)$")
        m = _RE_VARIANT.match(corrected) or _RE_VARIANT.match(s)
        if m:
            numeric_part = m.group(1)
            normalized = _normalize_code(numeric_part)
            return True, normalized  # corrigeable : on retire le suffixe

        # Non corrigeable : on retourne "" pour signaler
        return True, ""

    normalized = _normalize_code(corrected)
    return is_malformed, normalized


def _similar_codes(code: str, tco_codes: set[str], max_candidates: int = 2) -> list[str]:
    """Trouve les codes TCO les plus proches d'un code inconnu.

    Utilise la distance de Levenshtein simplifiée pour détecter les erreurs
    de frappe : chiffres transposés, segment manquant, etc.
    Retourne jusqu'à `max_candidates` codes similaires (distance ≤ 2).
    """
    if not code or not tco_codes:
        return []

    candidates = []
    for tco in tco_codes:
        # Filtrer rapidement sur le premier segment (même section)
        if code.split(".")[0] != tco.split(".")[0]:
            continue
        # Distance caractère naïve (insertion/suppression/substitution)
        dist = _levenshtein(code, tco)
        if dist <= 2:
            candidates.append((dist, tco))

    candidates.sort()
    return [c for _, c in candidates[:max_candidates]]


def _levenshtein(a: str, b: str) -> int:
    """Distance de Levenshtein entre deux chaînes."""
    if a == b:
        return 0
    if not a:
        return len(b)
    if not b:
        return len(a)
    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, 1):
        curr = [i]
        for j, cb in enumerate(b, 1):
            curr.append(min(prev[j] + 1, curr[j - 1] + 1, prev[j - 1] + (0 if ca == cb else 1)))
        prev = curr
    return prev[len(b)]


def _match_by_desig(
    dpgf_desig: str,
    tco_desig_index: "dict[str, tuple[str, int]]",
    threshold: float = 0.50,
) -> "tuple[str, int, float] | tuple[None, None, float]":
    """Trouve le meilleur match TCO par désignation (similarité Jaccard sur mots).

    Utilisé quand le code DPGF est vide : cherche un article TCO dont la
    désignation est suffisamment proche (≥ threshold) de celle du DPGF.

    Returns:
        (code_tco, df_idx, score) si match ≥ threshold
        (None, None, 0.0)         sinon
    """
    if not dpgf_desig or not tco_desig_index:
        return None, None, 0.0

    dpgf_words = set(dpgf_desig.lower().split())
    if len(dpgf_words) < 2:
        return None, None, 0.0

    best_score = 0.0
    best_code: str | None = None
    best_idx: int | None = None

    for tco_d, (code, idx) in tco_desig_index.items():
        tco_words = set(tco_d.split())
        if not tco_words:
            continue
        union = dpgf_words | tco_words
        inter = dpgf_words & tco_words
        score = len(inter) / len(union)
        if score > best_score:
            best_score = score
            best_code = code
            best_idx = idx

    if best_score >= threshold:
        return best_code, best_idx, best_score
    return None, None, 0.0


def _qc_check_dpgf_row(
    dpgf_row: "pd.Series",
    code: str,
    company_name: str,
    tco_desig: str,
    tco_codes: set[str],
    is_code_matched: bool,
) -> list[dict]:
    """Contrôle qualité complet d'une ligne DPGF.

    Vérifie :
    1. Erreur de calcul : Qu × Px_U_HT ≠ Px_Tot_HT (tolérance 1%)
    2. Désignation vide
    3. Prix unitaire manquant (Qu. renseigné mais Px_U_HT absent)
    4. Quantité nulle avec prix non nul
    5. Valeurs négatives (prix ou quantité)
    6. Texte dans colonne numérique (Qu. ou Prix)
    7. Code non trouvé dans le TCO → chercher un code proche
    8. Désignation très différente du TCO (similarité < 40%)

    Returns:
        Liste d'alertes {type, color, code, message}.
    """
    alerts: list[dict] = []
    row_type = dpgf_row.get("row_type", "article")
    if row_type not in ("article", "sub_section"):
        return alerts

    desig = str(dpgf_row.get("Désignation", "") or "").strip()
    raw_qu = dpgf_row.get("Qu.")
    raw_pu = dpgf_row.get("Px_U_HT")
    raw_tot = dpgf_row.get("Px_Tot_HT")

    # --- Helpers de conversion ---
    def _to_float(v: object) -> "float | None":
        if v is None:
            return None
        try:
            return float(v)
        except (ValueError, TypeError):
            return None

    def _is_text_in_num(v: object) -> bool:
        """True si la valeur est une chaîne non numérique (annotation dans champ numérique)."""
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return False
        if isinstance(v, (int, float)):
            return False
        s = str(v).strip()
        try:
            float(s.replace(",", "."))
            return False
        except ValueError:
            return bool(s)  # non vide ET non convertible

    qu = _to_float(raw_qu)
    pu = _to_float(raw_pu)
    tot = _to_float(raw_tot)

    # 6. Texte dans colonnes numériques
    if _is_text_in_num(raw_qu):
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Texte dans la colonne Quantité : '{raw_qu}' — annotation mal placée ?",
            }
        )
    if _is_text_in_num(raw_pu):
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Texte dans la colonne Px U. HT : '{raw_pu}' — annotation mal placée ?",
            }
        )
    if _is_text_in_num(raw_tot):
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Texte dans la colonne Px Tot HT : '{raw_tot}' — annotation mal placée ?",
            }
        )

    # 5. Valeurs négatives
    if qu is not None and qu < 0:
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Quantité négative : {qu}",
            }
        )
    if pu is not None and pu < 0:
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Prix unitaire négatif : {pu} €",
            }
        )
    if tot is not None and tot < 0:
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Prix total négatif : {tot} €",
            }
        )

    # 4. Quantité nulle avec prix non nul
    if qu is not None and qu == 0 and pu is not None and pu != 0:
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Quantité = 0 mais Px U. HT = {pu} € — ligne incohérente",
            }
        )

    # 1. Erreur de calcul Qu × Px_U_HT ≠ Px_Tot_HT (tolérance 1%)
    if qu is not None and pu is not None and tot is not None and qu != 0 and pu != 0 and tot != 0:
        expected = qu * pu
        if abs(expected) > 0.001:
            ratio = abs(tot - expected) / abs(expected)
            if ratio > 0.01:
                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": (
                            f"Erreur de calcul : {qu} × {pu} = {expected:.2f} "
                            f"mais Px Tot = {tot:.2f} (écart {ratio:.1%})"
                        ),
                    }
                )

    # 2. Désignation vide
    if not desig and code:
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": "Désignation vide pour ce poste",
            }
        )

    # 3. Prix unitaire absent alors que quantité renseignée
    if qu and qu != 0 and (pu is None or pu == 0) and (tot is None or tot == 0):
        alerts.append(
            {
                "type": "warning",
                "color": "orange",
                "code": code,
                "message": f"Quantité renseignée ({qu}) mais prix unitaire absent",
            }
        )

    # 7. Code inconnu → chercher un code proche (typo probable)
    if not is_code_matched and code:
        similar = _similar_codes(code, tco_codes)
        if similar:
            suggestions = ", ".join(f"'{s}'" for s in similar)
            alerts.append(
                {
                    "type": "error",
                    "color": "red",
                    "code": code,
                    "message": (
                        f"Code '{code}' absent du TCO — code(s) proche(s) : {suggestions} ?"
                    ),
                }
            )

    # 8. Désignation très différente de celle du TCO (si ligne matchée)
    if is_code_matched and desig and tco_desig:
        tco_d = tco_desig.strip().lower()
        dpgf_d = desig.lower()
        # Ratio de mots communs (jaccard sur tokens)
        tco_words = set(tco_d.split())
        dpgf_words = set(dpgf_d.split())
        if tco_words and dpgf_words:
            union = tco_words | dpgf_words
            inter = tco_words & dpgf_words
            jaccard = len(inter) / len(union)
            # Seulement si les deux désignations sont suffisamment longues pour être significatives
            if jaccard < 0.35 and len(tco_words) >= 3 and len(dpgf_words) >= 3:
                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": (
                            f"Désignation très différente du TCO "
                            f"(similarité {jaccard:.0%}) — vérifier le poste"
                        ),
                    }
                )

    return alerts


def _build_section_index(df: pd.DataFrame) -> dict[str, int]:
    """
    PERF-1 : Pré-indexe les section_headers par code pour lookup O(1).
    Retourne un dict {code: dataframe_index}.
    """
    return {
        _normalize_code(row["Code"]): idx
        for idx, row in df.iterrows()
        if row["row_type"] == "section_header" and row["Code"]
    }


def _get_children_total(
    df: pd.DataFrame,
    parent_code: str,
    total_col: str,
    children_index: dict[str, list[int]],
) -> Decimal:
    """
    Calcule la somme des valeurs d'une colonne pour tous les enfants
    (articles ET sub_sections) d'une section.
    """
    prefix = parent_code + "."
    total = Decimal("0.0")
    for code, idx_list in children_index.items():
        if not code.startswith(prefix):
            continue
        for idx in idx_list:
            row = df.loc[idx]
            if row["row_type"] in ("article", "sub_section"):
                val = row.get(total_col)
                if val is not None:
                    try:
                        if isinstance(val, Decimal):
                            total += val
                        else:
                            total += Decimal(str(val))
                    except (ValueError, TypeError):  # noqa: S110
                        pass
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


def _build_new_row(
    code: str,
    dpgf_row: "pd.Series",
    merged_df: "pd.DataFrame",
    col_u: str,
    col_qu: str,
    col_pu: str,
    col_tot: str,
    col_com: str,
    parent_code: str,
    original_row_idx: int,
) -> dict:
    """
    Construit le dict d'une nouvelle ligne article à insérer dans le TCO.
    Utilisé lors de l'insertion directe et du fallback hiérarchique (dédupliqué).
    """
    new_row: dict = {
        "Code": code,
        "Désignation": dpgf_row["Désignation"],
        "Qu.": None,
        "U": dpgf_row.get("U", ""),
        "Px_U_HT": None,
        "Px_Tot_HT": None,
        "Entete": dpgf_row.get("Entete", "Ouv_Art"),
        "row_type": "article",
        "original_row": original_row_idx,
        "parent_code": parent_code,
        "is_extra_line": True,
    }
    for col in merged_df.columns:
        if any(
            suffix in col for suffix in ["_U.", "_Qu.", "_Px_U_HT", "_Px_Tot_HT", "_Commentaire"]
        ):
            new_row[col] = None
    new_row[col_u] = dpgf_row.get("U", "")
    new_row[col_qu] = dpgf_row["Qu."]
    new_row[col_pu] = dpgf_row["Px_U_HT"]
    new_row[col_tot] = dpgf_row["Px_Tot_HT"]
    new_row[col_com] = dpgf_row.get("Commentaire", "")
    new_row["skip_sum"] = dpgf_row.get("skip_sum", False)
    return new_row


def _apply_total_lines(
    df: pd.DataFrame,
    total_col: str,
    montant_ht: Decimal,
    tva_rate: float,
    montant_options: Decimal | None = None,
) -> None:
    """
    Met à jour les lignes total_line (Montant HT / TVA / TTC).
    Extrait pour éviter la duplication entre compute_section_totals et
    _compute_ht_tva_ttc_base.
    """
    if montant_ht <= 0:
        for idx, row in df.iterrows():
            if row["row_type"] == "total_line":
                df.at[idx, total_col] = 0.0
        return

    d_tva_rate = Decimal(str(tva_rate))
    tva = (montant_ht * d_tva_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    montant_ttc = (montant_ht + tva).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    term_map = {
        "montant ht": montant_ht,
        "tva": tva,
        "ttc": montant_ttc,
        "option": montant_options if montant_options is not None else Decimal("0.0"),
    }

    for idx, row in df.iterrows():
        if row["row_type"] != "total_line":
            continue
        desig = str(row.get("Désignation", "")).strip().lower()
        col_com = total_col.replace("Px_Tot_HT", "Commentaire")

        # Priorité à la détection des options
        if "option" in desig or "variante" in desig:
            if montant_options is not None:
                df.at[idx, total_col] = float(montant_options)
            if col_com in df.columns:
                df.at[idx, col_com] = ""
            continue

        for key, val in term_map.items():
            if key in desig:
                df.at[idx, total_col] = float(val) if val is not None else 0.0
                if col_com in df.columns:
                    df.at[idx, col_com] = ""
                break


# ---------------------------------------------------------------------------
# Main merger
# ---------------------------------------------------------------------------


def merge_company_into_tco(
    tco_df: pd.DataFrame,
    dpgf_df: pd.DataFrame,
    company_name: str,
    tva_rate: float = TVA_DEFAULT,
    parse_alerts: list[dict] | None = None,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Fusionne un DPGF normalisé dans le TCO.

    Args:
        tco_df       : DataFrame du TCO modèle (de parse_tco)
        dpgf_df      : DataFrame normalisé du DPGF (de parse_dpgf)
        company_name : nom de l'entreprise
        tva_rate     : taux de TVA (paramétrable)
        parse_alerts : alertes du parser (pour extraire info_ht)

    Returns:
        merged_df : DataFrame avec colonnes entreprise ajoutées
        alerts    : liste d'alertes (codes non trouvés)
    """
    log.info("Fusion DPGF pour l'entreprise : %s (TVA=%.2f)", company_name, tva_rate)
    merged_df = tco_df.copy()
    alerts = []

    col_u = f"{company_name}_U."
    col_qu = f"{company_name}_Qu."
    col_pu = f"{company_name}_Px_U_HT"
    col_tot = f"{company_name}_Px_Tot_HT"
    col_com = f"{company_name}_Commentaire"

    merged_df[col_u] = None
    merged_df[col_qu] = None
    merged_df[col_pu] = None
    merged_df[col_tot] = 0.0
    merged_df[col_com] = None

    if dpgf_df.empty:
        log.warning("Le DPGF de %s est vide. Aucune fusion effectuée.", company_name)
        return merged_df, [
            {
                "type": "warning",
                "color": "orange",
                "row": 0,
                "code": "",
                "message": "Fichier DPGF vide ou non reconnu.",
            }
        ]

    # PERF-1 : index TCO codes → O(1) lookup
    tco_code_index: dict[str, int] = {}
    recap_by_parent: dict[str, int] = {}

    # Itération optimisée (100x plus rapide qu'iterrows simple)
    for idx, row_type, raw_code, parent_c in zip(
        merged_df.index,
        merged_df["row_type"],
        merged_df["Code"],
        merged_df["parent_code"],
        strict=False,
    ):
        code = _normalize_code(raw_code)
        if code and row_type not in ("empty", "recap", "recap_summary"):
            if code not in tco_code_index:
                tco_code_index[code] = idx

        if row_type == "recap":
            pc = _normalize_code(parent_c)
            if pc not in recap_by_parent:
                recap_by_parent[pc] = idx

    # Set de codes TCO pour le fuzzy matching
    tco_codes_set = set(tco_code_index.keys())

    # Index désignation TCO → (code, df_idx) pour matching par désignation (code vide)
    # Exclut les lignes ajoutées dynamiquement par des fusions précédentes (is_extra_line)
    tco_desig_index: dict[str, tuple[str, int]] = {}
    for idx, row in merged_df.iterrows():
        if row.get("is_extra_line"):
            continue
        if row["row_type"] not in ("article", "sub_section"):
            continue
        code = _normalize_code(row["Code"])
        desig = str(row.get("Désignation", "") or "").strip().lower()
        if code and desig:
            tco_desig_index[desig] = (code, idx)

    # Fusion articles + sub_sections du DPGF
    dpgf_data = dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])]
    matched_count = 0
    insertions: list[tuple[int, int, dict]] = []  # (target_idx, order, row_data)
    unclassified_std: list[dict] = []
    unclassified_opt: list[dict] = []
    unclassified_nocode: list[dict] = []  # Code vraiment vide → section "Articles sans code"
    unclassified_counter = 0
    # Codes TCO déjà remplis par cette entreprise : détecte les variantes (2.4.1.5b
    # corrigé en 2.4.1.5 alors que 2.4.1.5 a déjà été traité → extra row, pas écrasement)
    already_filled_codes: set[str] = set()

    # Remplacement complet du iterrows par to_dict('records')
    # Les dicts sont nettement plus rapides à parcourir et manipuler que les pd.Series
    dpgf_records = dpgf_data.to_dict("records")
    for dpgf_row in dpgf_records:
        raw_code = dpgf_row.get("Code", "")
        code = _normalize_code(raw_code)

        # --- Détection code malformé ---
        is_malformed, corrected_code = _detect_malformed_code(raw_code)
        if is_malformed:
            raw_str = str(raw_code).strip()
            if corrected_code:
                # Correction possible → utiliser le code corrigé
                # L'alerte sera émise après le matching (warning si ok, pas error)
                log.warning(
                    "Code DPGF malformé corrigé : %r → %r (%s)",
                    raw_str,
                    corrected_code,
                    company_name,
                )
                code = corrected_code
            else:
                # Non corrigeable → error immédiate (ira en section 99)
                log.warning(
                    "Code DPGF non corrigeable : %r (%s) — insertion en extra",
                    raw_str,
                    company_name,
                )
                code = ""  # force le bloc "code non trouvé" plus bas
                alerts.append(
                    {
                        "type": "error",
                        "color": "red",
                        "code": "",
                        "message": (
                            f"Code incorrect dans le DPGF {company_name}"
                            f" : '{raw_str}' (non corrigeable)"
                        ),
                    }
                )

        # --- GESTION DES ARTICLES NON CLASSABLES ---
        if not code or code not in tco_code_index:
            found_insertion = False
            parent_code = "99"  # Default base parent

            # --- CODE VIDE : tentative de correspondance par désignation ---
            # Exclus : lignes option (is_option=True) → jamais matcher sur un article TCO
            # de base, elles doivent aller en OPT_DYN pour ne pas écraser les quantités.
            if not code and not dpgf_row.get("is_option"):
                dpgf_desig_raw = str(dpgf_row.get("Désignation", "")).strip()
                m_code, m_idx, m_score = _match_by_desig(dpgf_desig_raw, tco_desig_index)
                if m_code is not None:
                    base_com = str(dpgf_row.get("Commentaire", "") or "").strip()
                    desig_note = (
                        f"⚠️ Code vide — correspondance désignation "
                        f"(code TCO : {m_code}, similarité {m_score:.0%})"
                    )
                    merged_df.at[m_idx, col_u] = dpgf_row.get("U", "")
                    merged_df.at[m_idx, col_qu] = dpgf_row["Qu."]
                    merged_df.at[m_idx, col_pu] = dpgf_row["Px_U_HT"]
                    merged_df.at[m_idx, col_tot] = dpgf_row["Px_Tot_HT"]
                    merged_df.at[m_idx, col_com] = (
                        f"{desig_note} ; {base_com}".strip(" ;") if base_com else desig_note
                    )
                    if dpgf_row.get("is_option"):
                        merged_df.at[m_idx, "is_option"] = True
                    matched_count += 1
                    found_insertion = True
                    log.info(
                        "Code vide ligne %s → match désignation code=%s (sim. %.0f%%)",
                        dpgf_row.get("original_row", "?"),
                        m_code,
                        m_score * 100,
                    )
                    alerts.append(
                        {
                            "type": "warning",
                            "color": "orange",
                            "code": m_code,
                            "message": (
                                f"Code vide → correspondance désignation "
                                f"code TCO : {m_code} (sim. {m_score:.0%}) — {company_name}"
                            ),
                        }
                    )

            # --- TENTATIVE D'INSERTION HIÉRARCHIQUE SI CODE EXISTE ---
            if not found_insertion and code:
                # Recherche O(1) du parent pour insertion
                parent_suggested = ".".join(code.split(".")[:-1])
                norm_parent = _normalize_code(parent_suggested)
                _added_note = "⚠️ Ligne ajoutée par l'entreprise (code absent du TCO de base)"
                if norm_parent in recap_by_parent:
                    idx = recap_by_parent[norm_parent]
                    parent_code = parent_suggested
                    new_row = _build_new_row(
                        code,
                        dpgf_row,
                        merged_df,
                        col_u,
                        col_qu,
                        col_pu,
                        col_tot,
                        col_com,
                        parent_code,
                        int(idx),
                    )
                    new_row["is_added"] = True
                    _cur_com = str(new_row.get(col_com) or "").strip()
                    new_row[col_com] = (
                        f"{_added_note} ; {_cur_com}".strip(" ;") if _cur_com else _added_note
                    )
                    insertions.append((int(idx), len(insertions), new_row))
                    matched_count += 1
                    found_insertion = True
                    alerts.append(
                        {
                            "type": "info",
                            "color": "blue",
                            "code": code,
                            "message": f"Code ajouté par l'entreprise : '{raw_code}' (inséré sous '{parent_suggested}') — {company_name}",
                        }
                    )
                else:
                    # Fallback hiérarchique ancêtre
                    fallback_parts = parent_suggested.split(".")
                    while len(fallback_parts) > 0 and not found_insertion:
                        fallback_parts.pop()
                        if not fallback_parts:
                            break
                        fb_parent = ".".join(fallback_parts)
                        norm_fb = _normalize_code(fb_parent)
                        if norm_fb in recap_by_parent:
                            idx = recap_by_parent[norm_fb]
                            parent_code = fb_parent
                            new_row = _build_new_row(
                                code,
                                dpgf_row,
                                merged_df,
                                col_u,
                                col_qu,
                                col_pu,
                                col_tot,
                                col_com,
                                parent_code,
                                int(idx),
                            )
                            new_row["is_added"] = True
                            _cur_com = str(new_row.get(col_com) or "").strip()
                            new_row[col_com] = (
                                f"{_added_note} ; {_cur_com}".strip(" ;")
                                if _cur_com
                                else _added_note
                            )
                            insertions.append((int(idx), len(insertions), new_row))
                            matched_count += 1
                            found_insertion = True
                            alerts.append(
                                {
                                    "type": "info",
                                    "color": "blue",
                                    "code": code,
                                    "message": f"Code ajouté par l'entreprise : '{raw_code}' (inséré sous '{fb_parent}') — {company_name}",
                                }
                            )
                            break

            if not found_insertion:
                unclassified_counter += 1

                # --- ÉVALUATION DIAGNOSTIC & PROTECTION DOUBLONS ---
                is_genuinely_nocode = not str(raw_code or "").strip()
                if is_genuinely_nocode:
                    reason = "Ligne sans code dans le DPGF"
                elif is_malformed:
                    reason = f"Code malformé non corrigeable ('{raw_code}')"
                else:
                    reason = f"Code '{raw_code}' inconnu"

                desig = str(dpgf_row.get("Désignation", ""))
                is_junk = bool(_RE_JUNK_TOTAL.search(desig))
                if is_junk:
                    reason += " | ⚠️ TOTAL/SOUS-TOTAL DÉTECTÉ (Exclu du TCO pour éviter doublon)"
                    dpgf_row["skip_sum"] = True

                orig_row_val = dpgf_row.get("original_row")
                orig_row_int = int(float(orig_row_val or 0)) if pd.notna(orig_row_val) else 0
                new_row = _build_new_row(
                    raw_code,
                    dpgf_row,
                    merged_df,
                    col_u,
                    col_qu,
                    col_pu,
                    col_tot,
                    col_com,
                    "99",
                    orig_row_int,
                )

                # Insertion du diagnostic dans le commentaire entreprise
                current_com = str(new_row.get(col_com) or "").strip()
                new_row[col_com] = f"{reason} ; {current_com}".strip(" ;")

                if dpgf_row.get("is_option"):
                    new_row["parent_code"] = "OPT_DYN"
                    unclassified_opt.append(new_row)
                elif is_genuinely_nocode:
                    # Ligne vraiment sans code → section dédiée "Articles sans code"
                    new_row["parent_code"] = "SANS_CODE"
                    unclassified_nocode.append(new_row)
                else:
                    unclassified_std.append(new_row)

                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": f"{reason} ({company_name})",
                    }
                )

        else:
            # --- MATCHING STANDARD ---
            idx = tco_code_index[code]
            raw_str = str(raw_code).strip()
            base_com = str(dpgf_row.get("Commentaire", "") or "").strip()

            is_duplicate = code in already_filled_codes
            if is_duplicate:
                # --- CAS DUPLICATA/VARIANTE : code déjà rempli par cette entreprise ---
                # On n'écrase pas la ligne TCO existante -> extra row insérée juste après
                if is_malformed and corrected_code:
                    var_note = (
                        f"⚠️ Variante de '{code}' "
                        f"(code DPGF original incorrect : '{raw_str}' — suffixe supprimé)"
                    )
                else:
                    var_note = f"⚠️ Duplicata de '{code}' (présent plusieurs fois dans le DPGF)"
                new_row = _build_new_row(
                    raw_str,
                    dpgf_row,
                    merged_df,
                    col_u,
                    col_qu,
                    col_pu,
                    col_tot,
                    col_com,
                    "",
                    int(idx),
                )
                new_row["Code"] = raw_str  # affiche le code variante d'origine
                new_row[col_com] = f"{var_note} ; {base_com}".strip(" ;") if base_com else var_note
                if dpgf_row.get("is_option"):
                    new_row["is_option"] = True
                insertions.append((int(idx), len(insertions), new_row))
                log.info(
                    "Variante '%s' → extra row après '%s' (%s)",
                    raw_str,
                    code,
                    company_name,
                )
                if is_malformed and corrected_code:
                    message = (
                        f"Code variante/duplicata '{raw_str}' → inséré après '{code}' "
                        f"(déjà rempli) — {company_name}"
                    )
                else:
                    message = (
                        f"Code variante '{raw_str}' → inséré après '{code}' "
                        f"(code de base déjà rempli) — {company_name}"
                    )

                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": message,
                    }
                )
                matched_count += 1

            else:
                # --- MATCH NORMAL (ou variante sans doublon) ---
                merged_df.at[idx, col_u] = dpgf_row.get("U", "")
                merged_df.at[idx, col_qu] = dpgf_row["Qu."]
                merged_df.at[idx, col_pu] = dpgf_row["Px_U_HT"]
                merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
                if is_malformed and corrected_code:
                    malform_note = f"⚠️ Code incorrect : '{raw_str}' — corrigé en '{corrected_code}'"
                    merged_df.at[idx, col_com] = (
                        f"{malform_note} ; {base_com}".strip(" ;") if base_com else malform_note
                    )
                    # Warning (pas error) : la correction a réussi
                    alerts.append(
                        {
                            "type": "warning",
                            "color": "orange",
                            "code": corrected_code,
                            "message": (
                                f"Code '{raw_str}' corrigé en '{corrected_code}' — {company_name}"
                            ),
                        }
                    )
                else:
                    merged_df.at[idx, col_com] = base_com

                if dpgf_row.get("is_option"):
                    merged_df.at[idx, "is_option"] = True

                already_filled_codes.add(code)

                # --- QC CHECK ---
                tco_desig = str(merged_df.at[idx, "Désignation"] or "").strip()
                alerts.extend(
                    _qc_check_dpgf_row(
                        dpgf_row, code, company_name, tco_desig, tco_codes_set, is_code_matched=True
                    )
                )
                matched_count += 1

    # P2 FIX : insertion O(n) en un seul passage
    if insertions or unclassified_std or unclassified_opt or unclassified_nocode:
        ins_by_pos: dict[int, list[tuple[int, dict]]] = defaultdict(list)
        for target_idx, order, row_data in insertions:
            ins_by_pos[target_idx].append((order, row_data))
        for pos in ins_by_pos:
            ins_by_pos[pos].sort(key=lambda x: x[0])

        new_rows: list[dict] = []
        for pos, record in enumerate(merged_df.to_dict("records")):
            if pos in ins_by_pos:
                for _, row_data in ins_by_pos[pos]:
                    new_rows.append(row_data)
            new_rows.append(record)

        # --- SECTIONS DYNAMIQUES À LA TOUTE FIN ---
        # 1. Section Options
        if unclassified_opt:
            if not any(r.get("Code") == "OPT_DYN" for r in new_rows):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "OPT_DYN",
                        "Désignation": "Options & Variantes (Hors-Bordereau)",
                        "Entete": "Bd_OPT_Bord",
                        "row_type": "section_header",
                        "parent_code": "",
                        "is_extra_line": True,
                        "is_option": True,
                    }
                )
            new_rows.extend(unclassified_opt)
            if not any(
                r.get("Code") == "OPT_DYN" and r.get("row_type") == "recap" for r in new_rows
            ):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "OPT_DYN",
                        "Désignation": "Total Options & Variantes",
                        "Entete": "Bord_OPT_Recap",
                        "row_type": "recap",
                        "parent_code": "OPT_DYN",
                        "is_extra_line": True,
                        "is_option": True,
                    }
                )
            if not any(
                r.get("Code") == "OPT_DYN" and r.get("row_type") == "recap_summary"
                for r in new_rows
            ):
                summary_row = {c: None for c in merged_df.columns}
                summary_row.update(
                    {
                        "Code": "OPT_DYN",
                        "Désignation": "Options & Variantes (Hors-Bordereau)",
                        "row_type": "recap_summary",
                        "is_extra_line": True,
                        "is_option": True,
                    }
                )
                insert_idx = next(
                    (i for i, r in enumerate(new_rows) if r.get("row_type") == "total_line"),
                    len(new_rows),
                )
                new_rows.insert(insert_idx, summary_row)

        # 2. Section 99
        if unclassified_std:
            if not any(r.get("Code") == "99" for r in new_rows):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "99",
                        "Désignation": "99 - Articles non classables (Code absent ou inconnu)",
                        "Entete": "Bd_99_Bord",
                        "row_type": "section_header",
                        "parent_code": "",
                        "is_extra_line": True,
                    }
                )
            new_rows.extend(unclassified_std)
            if not any(r.get("Code") == "99" and r.get("row_type") == "recap" for r in new_rows):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "99",
                        "Désignation": "Total Articles non classables",
                        "Entete": "Bord_99_Recap",
                        "row_type": "recap",
                        "parent_code": "99",
                        "is_extra_line": True,
                    }
                )
            if not any(
                r.get("Code") == "99" and r.get("row_type") == "recap_summary" for r in new_rows
            ):
                summary_row = {c: None for c in merged_df.columns}
                summary_row.update(
                    {
                        "Code": "99",
                        "Désignation": "99 - Articles non classables (Code absent ou inconnu)",
                        "row_type": "recap_summary",
                        "is_extra_line": True,
                        "is_option": False,
                    }
                )
                insert_idx = next(
                    (i for i, r in enumerate(new_rows) if r.get("row_type") == "total_line"),
                    len(new_rows),
                )
                new_rows.insert(insert_idx, summary_row)

        # 3. Section SANS_CODE (Articles présents sans code dans le DPGF)
        if unclassified_nocode:
            if not any(r.get("Code") == "SANS_CODE" for r in new_rows):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "SANS_CODE",
                        "Désignation": "Articles sans code (présents dans le DPGF mais sans code)",
                        "Entete": "Bd_SC_Bord",
                        "row_type": "section_header",
                        "parent_code": "",
                        "is_extra_line": True,
                    }
                )
            new_rows.extend(unclassified_nocode)
            if not any(
                r.get("Code") == "SANS_CODE" and r.get("row_type") == "recap" for r in new_rows
            ):
                new_rows.append({c: None for c in merged_df.columns})
                new_rows[-1].update(
                    {
                        "Code": "SANS_CODE",
                        "Désignation": "Total Articles sans code",
                        "Entete": "Bord_SC_Recap",
                        "row_type": "recap",
                        "parent_code": "SANS_CODE",
                        "is_extra_line": True,
                    }
                )
            if not any(
                r.get("Code") == "SANS_CODE" and r.get("row_type") == "recap_summary"
                for r in new_rows
            ):
                summary_row = {c: None for c in merged_df.columns}
                summary_row.update(
                    {
                        "Code": "SANS_CODE",
                        "Désignation": "Articles sans code",
                        "row_type": "recap_summary",
                        "is_extra_line": True,
                        "is_option": False,
                    }
                )
                insert_idx = next(
                    (i for i, r in enumerate(new_rows) if r.get("row_type") == "total_line"),
                    len(new_rows),
                )
                new_rows.insert(insert_idx, summary_row)

        merged_df = pd.DataFrame(new_rows).reset_index(drop=True)

    # Commentaire sur les lignes Montant HT/TVA/TTC si des articles SANS_CODE existent (Feature 3)
    if unclassified_nocode:
        nocode_total = sum(
            float(r.get(col_tot) or 0) for r in unclassified_nocode if pd.notna(r.get(col_tot))
        )
        if nocode_total != 0:
            nocode_count = len(unclassified_nocode)
            note = (
                f"⚠️ {nocode_count} article(s) sans code non inclus"
                f" dans ce total (montant HT non classé :"
                f" {nocode_total:,.2f} €)"
                f" — voir section 'Articles sans code'"
            )
            for i in merged_df.index[merged_df["row_type"] == "total_line"]:
                cur = str(merged_df.at[i, col_com] or "").strip()
                merged_df.at[i, col_com] = f"{note} ; {cur}".strip(" ;") if cur else note

    # Taux de correspondance
    total_dpgf = len(dpgf_data)
    if total_dpgf > 0:
        match_rate = matched_count / total_dpgf * 100
        unmatched = total_dpgf - matched_count
        if match_rate < 50:
            msg = f"DPGF ignoré — trop peu de codes correspondent au template ({matched_count}/{total_dpgf} codes matchés, soit {match_rate:.0f}%)."
            log.error(
                "Match rate critique %s : %.1f%% (%d/%d)",
                company_name,
                match_rate,
                matched_count,
                total_dpgf,
            )
            alerts.insert(
                0, {"type": "error", "color": "red", "row": 0, "code": "", "message": msg}
            )
            return tco_df, alerts
        elif match_rate < 90:
            msg = f"Correspondance DPGF/Template partielle : {match_rate:.0f}% — {unmatched} codes non trouvés sur {total_dpgf}"
            log.warning(msg)
            alerts.insert(
                0, {"type": "warning", "color": "orange", "row": 0, "code": "", "message": msg}
            )
        log.info(
            "Match rate %s : %.1f%% (%d/%d)", company_name, match_rate, matched_count, total_dpgf
        )

    log.info("Fusion terminée : %d lignes matchées, %d non trouvées", matched_count, len(alerts))

    compute_section_totals(merged_df, col_tot, tva_rate=tva_rate)

    # Vérification de l'écart HT
    if parse_alerts:
        extracted_ht = None
        for a in parse_alerts:
            if a.get("type") == "info_ht":
                extracted_ht = a.get("value")
                break

        if extracted_ht is not None:
            tco_ht = 0.0
            for idx in merged_df.index[merged_df["row_type"] == "total_line"]:
                if "montant ht" in str(merged_df.at[idx, "Désignation"]).lower():
                    val = merged_df.at[idx, col_tot]
                    try:
                        tco_ht = float(val) if val is not None else 0.0
                    except (ValueError, TypeError):
                        tco_ht = 0.0
                    break

            ecart = abs(tco_ht - extracted_ht)
            if ecart > 1.0:
                msg_alert = f"Le Montant HT déclaré ({extracted_ht:,.2f} €) diffère du calcul TCO ({tco_ht:,.2f} €) — Écart : {ecart:,.2f} €"
                log.warning(
                    "Écart HT pour %s: déclaré=%.2f, calculé=%.2f (écart %.2f)",
                    company_name,
                    extracted_ht,
                    tco_ht,
                    ecart,
                )

                alerts.insert(
                    0,
                    {
                        "type": "warning",
                        "color": "orange",
                        "row": 0,
                        "code": "",
                        "message": msg_alert,
                    },
                )

                for idx in merged_df.index[merged_df["row_type"] == "total_line"]:
                    if "montant ht" in str(merged_df.at[idx, "Désignation"]).lower():
                        cur = str(merged_df.at[idx, col_com] or "").strip()
                        merged_df.at[idx, col_com] = f"⚠️ {msg_alert} ; {cur}".strip(" ;")
                        break

    return merged_df, alerts


# ---------------------------------------------------------------------------
# Fusion de toutes les entreprises (logique métier extraite de l'UI)
# ---------------------------------------------------------------------------


def merge_all_companies(
    tco_df: pd.DataFrame,
    company_data: dict,
    tva_rate: float = TVA_DEFAULT,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Fusionne toutes les entreprises dans le TCO de base.

    Args:
        tco_df       : DataFrame du TCO modèle de base (parse_tco)
        company_data : dict {nom: {"dpgf_df": df, "parse_alerts": [...], "filename": str}}
        tva_rate     : taux de TVA

    Returns:
        merged_df  : DataFrame avec toutes les colonnes entreprise
        all_alerts : toutes les alertes (parse + merge), taguées par entreprise
    """
    log.info(
        "Reconstruction TCO. TVA=%.2f. Entreprises=%s",
        tva_rate,
        list(company_data.keys()),
    )
    merged: pd.DataFrame = tco_df.copy()
    all_alerts: list[dict] = []

    for comp_name, comp_data in company_data.items():
        merged, merge_alerts = merge_company_into_tco(
            merged,
            comp_data["dpgf_df"],
            comp_name,
            tva_rate=tva_rate,
            parse_alerts=comp_data.get("parse_alerts", []),
        )
        for alert in comp_data.get("parse_alerts", []):
            alert["company"] = comp_name
        for alert in merge_alerts:
            alert["company"] = comp_name
        all_alerts.extend(comp_data.get("parse_alerts", []))
        all_alerts.extend(merge_alerts)

    # CQ : Vérification de la cohérence des unités entre entreprises
    unit_alerts = _check_units_consistency(merged)
    all_alerts.extend(unit_alerts)

    return merged, all_alerts


def _check_units_consistency(df: pd.DataFrame) -> list[dict]:
    """
    Vérifie que toutes les entreprises utilisent la même unité pour un article donné.
    Génère une alerte orange si des différences sont détectées (hors estimation).
    """
    alerts: list[dict] = []
    # Identifier les colonnes d'unités entreprise (se terminant par _U.)
    unit_cols = [c for c in df.columns if c.endswith("_U.")]
    if len(unit_cols) < 2:
        return alerts

    art_df = df[df["row_type"] == "article"]
    for idx, row in art_df.iterrows():
        # Collecter les unités renseignées (non vides)
        units = {}
        for col in unit_cols:
            u = str(row.get(col) or "").strip()
            if u:
                company = col.replace("_U.", "")
                units[company] = u

        if len(set(units.values())) > 1:
            # Incohérence détectée
            detail = ", ".join([f"{c}: {u}" for c, u in units.items()])
            alerts.append(
                {
                    "type": "warning",
                    "color": "orange",
                    "code": row.get("Code", ""),
                    "message": f"Unités hétérogènes entre entreprises : {detail}",
                    "company": "CONTRÔLE QUALITÉ",
                }
            )
    return alerts


# ---------------------------------------------------------------------------
# Calcul des sous-totaux
# ---------------------------------------------------------------------------


def compute_section_totals(
    df: pd.DataFrame,
    total_col: str,
    tva_rate: float = TVA_DEFAULT,
) -> None:
    """
    Recalcule les totaux pour :
      1. section_header (01.X) : somme directe des articles + sub_sections
      2. recap (Code vide)     : reçoit le total de sa section parente
      3. recap_summary         : reçoit le total de leur section
      4. total_line            : Montant HT / TVA / TTC

    Args:
        total_col : colonne à calculer (ex: "MAB SUD-OUEST_Px_Tot_HT")
        tva_rate  : taux de TVA (paramétrable)
    """
    section_header_index = _build_section_index(df)

    # Forcer la colonne en object pour accepter Decimal sans erreur
    if total_col in df.columns:
        df[total_col] = df[total_col].astype(object)

    # PERF-7 : Précalcul des sommes par préfixe parent en un seul passage O(N*depth)
    # Évite O(S×N) itérations (une par section_header) de l'ancien _get_children_total.
    parent_sums: dict[str, Decimal] = defaultdict(Decimal)
    for _idx, row in df.iterrows():
        if row["row_type"] not in ("article", "sub_section") or row.get("skip_sum"):
            continue
        val = row.get(total_col)
        if val is None or pd.isna(val):
            continue
        try:
            v = val if isinstance(val, Decimal) else Decimal(str(val))
            if v.is_nan():
                continue
        except (ValueError, TypeError, InvalidOperation, ArithmeticError):  # noqa: S110
            continue

        # Propager via parent_code explicite (prioritaire pour sections dynamiques)
        p_code = _normalize_code(row.get("parent_code"))
        if p_code:
            parent_sums[p_code] += v

        # Propager via hiérarchie des codes (ex: 01.1.2 -> 01.1 -> 01)
        code = _normalize_code(row.get("Code"))
        if code:
            parts = code.split(".")
            for i in range(1, len(parts)):
                parent_sums[".".join(parts[:i])] += v

    # Passe 1 : totaux des section_headers (lookup O(1))
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or row["row_type"] != "section_header":
            continue
        if code in parent_sums:
            total = parent_sums[code].quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            df.at[idx, total_col] = float(total)

    # Passe 2 : propager vers les lignes recap
    for idx, row in df.iterrows():
        if row["row_type"] != "recap":
            continue
        parent = _normalize_code(row.get("parent_code", ""))
        if parent and parent in section_header_index:
            s_idx = section_header_index[parent]
            val = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val

    # Passe 3 : recap_summary
    for idx, row in df.iterrows():
        if row["row_type"] != "recap_summary":
            continue
        code = _normalize_code(row["Code"])
        if code and code in section_header_index:
            s_idx = section_header_index[code]
            val = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val
            # Propager le flag is_option au recap_summary
            if "is_option" in df.columns and df.at[s_idx, "is_option"]:
                df.at[idx, "is_option"] = True

    # Passe 3b : vider les section_headers — le total est désormais porté
    # uniquement par la ligne recap (évite le doublon visuel section_header / recap).
    # Placé après Passe 3 car recap_summary lit encore la valeur depuis section_header.
    for idx, row in df.iterrows():
        if row["row_type"] == "section_header":
            df.at[idx, total_col] = None

    # Passe 4 : Montant HT / TVA / TTC — somme des recap_summary (= lignes du récapitulatif)
    # Les recap_summary ont été remplis en Passe 3 depuis leur section_header,
    # donc leur somme reflète exactement ce qui est affiché dans le récapitulatif.
    # ON EXCLUT les options du Montant HT principal.
    montant_ht = Decimal("0.0")
    montant_options = Decimal("0.0")

    for _idx, row in df.iterrows():
        if row["row_type"] == "recap_summary":
            val = row.get(total_col)
            if val is not None and not pd.isna(val):
                try:
                    v = val if isinstance(val, Decimal) else Decimal(str(val))
                    if not v.is_nan():
                        if "is_option" in row.index and row.get("is_option"):
                            montant_options += v
                        else:
                            montant_ht += v
                except (ValueError, TypeError, InvalidOperation):  # noqa: S110
                    pass

    _apply_total_lines(df, total_col, montant_ht, tva_rate, montant_options=montant_options)

    # Colonne de base Px_Tot_HT (si on vient de calculer une colonne entreprise)
    if total_col != "Px_Tot_HT":
        _compute_ht_tva_ttc_base(df, tva_rate)


def _compute_ht_tva_ttc_base(df: pd.DataFrame, tva_rate: float = TVA_DEFAULT) -> None:
    """Calcule Montant HT, TVA, TTC pour la colonne de base Px_Tot_HT."""
    if "Px_Tot_HT" not in df.columns:
        return
    # Exclure les options (is_option=True) du montant HT de base — cohérent avec
    # compute_section_totals qui fait déjà cette exclusion pour les colonnes entreprise.
    is_option = df.get("is_option", pd.Series(False, index=df.index)).fillna(False).astype(bool)
    mask = (df["row_type"] == "recap_summary") & df["Px_Tot_HT"].notna() & ~is_option
    montant_ht = Decimal("0.0")
    for val in df.loc[mask, "Px_Tot_HT"]:
        if pd.isna(val):
            continue
        try:
            v = val if isinstance(val, Decimal) else Decimal(str(val))
            if not v.is_nan():
                montant_ht += v
        except (ValueError, TypeError, InvalidOperation):  # noqa: S110
            pass

    _apply_total_lines(df, "Px_Tot_HT", montant_ht, tva_rate)
