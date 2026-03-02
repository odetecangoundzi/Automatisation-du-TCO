"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Calcule les sous-totaux en se basant sur la hiérarchie Code (01.X → 01.X.Y...).
Les lignes "recap" (Code vide, Entete Bord_xxx_Recap) reçoivent le total
de leur section parente.
"""

from collections import defaultdict
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd

from config import TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)


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
    - Des caractères non alphanumériques  ("2.1.3-a", "2.1.3/")
    - Des points doublés ou finaux        ("2..1", "2.1.")

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
        # Non corrigeable : on retourne la chaîne brute pour la signaler
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
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Texte dans la colonne Quantité : '{raw_qu}' — annotation mal placée ?",
            }
        )
    if _is_text_in_num(raw_pu):
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": code,
                "message": f"Texte dans la colonne Px U. HT : '{raw_pu}' — annotation mal placée ?",
            }
        )
    if _is_text_in_num(raw_tot):
        alerts.append(
            {
                "type": "error",
                "color": "red",
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
    if qu is not None and pu is not None and tot is not None and qu != 0 and pu != 0:
        expected = qu * pu
        if abs(expected) > 0.001:
            ratio = abs(tot - expected) / abs(expected)
            if ratio > 0.01:
                alerts.append(
                    {
                        "type": "error",
                        "color": "red",
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
    new_row[col_u] = dpgf_row.get("U.", "")
    new_row[col_qu] = dpgf_row["Qu."]
    new_row[col_pu] = dpgf_row["Px_U_HT"]
    new_row[col_tot] = dpgf_row["Px_Tot_HT"]
    new_row[col_com] = dpgf_row.get("Commentaire", "")
    return new_row


def _apply_total_lines(
    df: pd.DataFrame,
    total_col: str,
    montant_ht: Decimal,
    tva_rate: float,
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
    term_map = {"montant ht": montant_ht, "tva": tva, "ttc": montant_ttc}

    for idx, row in df.iterrows():
        if row["row_type"] != "total_line":
            continue
        desig = str(row.get("Désignation", "")).strip().lower()
        for key, val in term_map.items():
            if key in desig:
                df.at[idx, total_col] = val
                break


# ---------------------------------------------------------------------------
# Main merger
# ---------------------------------------------------------------------------


def merge_company_into_tco(
    tco_df: pd.DataFrame,
    dpgf_df: pd.DataFrame,
    company_name: str,
    tva_rate: float = TVA_DEFAULT,
) -> tuple[pd.DataFrame, list[dict]]:
    """
    Fusionne un DPGF normalisé dans le TCO.

    Args:
        tco_df       : DataFrame du TCO modèle (de parse_tco)
        dpgf_df      : DataFrame normalisé du DPGF (de parse_dpgf)
        company_name : nom de l'entreprise
        tva_rate     : taux de TVA (paramétrable)

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
    merged_df[col_tot] = None
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
    for idx, row in merged_df.iterrows():
        code = _normalize_code(row["Code"])
        if code and row["row_type"] not in ("empty", "recap", "recap_summary"):
            if code not in tco_code_index:
                tco_code_index[code] = idx

    # Set de codes TCO pour le fuzzy matching
    tco_codes_set = set(tco_code_index.keys())

    # Fusion articles + sub_sections du DPGF
    dpgf_data = dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])]
    matched_count = 0
    insertions: list[tuple[int, int, dict]] = []  # (target_idx, order, row_data)

    # --- Détection de codes en doublon dans le DPGF ---
    dpgf_codes_all = [
        _normalize_code(r["Code"]) for _, r in dpgf_data.iterrows() if _normalize_code(r["Code"])
    ]
    seen_codes: set[str] = set()
    duplicate_codes: set[str] = set()
    for c in dpgf_codes_all:
        if c in seen_codes:
            duplicate_codes.add(c)
        seen_codes.add(c)
    for dup in sorted(duplicate_codes):
        alerts.append(
            {
                "type": "error",
                "color": "red",
                "code": dup,
                "message": f"Code '{dup}' en doublon dans le DPGF {company_name} — seule la première occurrence est conservée",
            }
        )

    for _, dpgf_row in dpgf_data.iterrows():
        raw_code = dpgf_row["Code"]
        code = _normalize_code(raw_code)

        # --- Détection code malformé ---
        is_malformed, corrected_code = _detect_malformed_code(raw_code)
        if is_malformed:
            raw_str = str(raw_code).strip()
            if corrected_code:
                # Correction possible → utiliser le code corrigé, signaler l'erreur
                log.warning(
                    "Code DPGF malformé corrigé : %r → %r (%s)",
                    raw_str,
                    corrected_code,
                    company_name,
                )
                code = corrected_code
            else:
                # Non corrigeable → signaler et forcer l'insertion comme extra
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
                    "code": corrected_code or "",
                    "message": (
                        f"Code incorrect dans le DPGF {company_name} : "
                        f"'{raw_str}'"
                        + (
                            f" → corrigé en '{corrected_code}'"
                            if corrected_code
                            else " (non corrigeable)"
                        )
                    ),
                }
            )

        if not code:
            px_tot = dpgf_row.get("Px_Tot_HT")
            try:
                if px_tot and Decimal(str(px_tot)) > 0:
                    desig = str(dpgf_row.get("Désignation", ""))[:60]
                    alerts.append(
                        {
                            "type": "warning",
                            "color": "orange",
                            "code": "",
                            "message": f"Ligne sans code ignorée (montant={px_tot} €) — {desig}",
                        }
                    )
            except Exception:  # noqa: S110
                pass
            continue

        if code in tco_code_index:
            idx = tco_code_index[code]
            merged_df.at[idx, col_u] = dpgf_row.get("U.", "")
            merged_df.at[idx, col_qu] = dpgf_row["Qu."]
            merged_df.at[idx, col_pu] = dpgf_row["Px_U_HT"]
            merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
            merged_df.at[idx, col_com] = dpgf_row.get("Commentaire", "")

            # --- DETECT MISMATCH IN QUANTITY & UNIT ---
            if dpgf_row["row_type"] == "article":
                tco_qu = merged_df.at[idx, "Qu."]
                tco_u = str(merged_df.at[idx, "U"] or "").strip().lower()
                dpgf_qu = dpgf_row["Qu."]
                dpgf_u = str(dpgf_row.get("U", "") or "").strip().lower()

                # Check Quantity mismatch
                try:
                    tco_qu_val = float(tco_qu) if tco_qu is not None else 0.0
                    dpgf_qu_val = float(dpgf_qu) if dpgf_qu is not None else 0.0
                    if tco_qu_val != dpgf_qu_val and tco_qu_val > 0:
                        alerts.append(
                            {
                                "type": "warning",
                                "color": "orange",
                                "code": code,
                                "message": f"Quantité divergente par rapport au modèle ({tco_qu} vs {dpgf_qu})",
                            }
                        )
                except (ValueError, TypeError):
                    pass

                # Check Unit mismatch — signalé comme ERREUR (rouge)
                if tco_u and dpgf_u and tco_u != dpgf_u:
                    alerts.append(
                        {
                            "type": "error",
                            "color": "red",
                            "code": code,
                            "message": (
                                f"Unité différente de l'estimation : "
                                f"estim.='{tco_u.upper()}' vs {company_name}='{dpgf_u.upper()}'"
                            ),
                        }
                    )

            # --- FULL QC CHECK ---
            tco_desig = str(merged_df.at[idx, "Désignation"] or "").strip()
            alerts.extend(
                _qc_check_dpgf_row(
                    dpgf_row, code, company_name, tco_desig, tco_codes_set, is_code_matched=True
                )
            )

            matched_count += 1
        else:
            parent_code = ".".join(code.split(".")[:-1])
            found_insertion = False

            # --- QC CHECK pour ligne non matchée (code inconnu) ---
            alerts.extend(
                _qc_check_dpgf_row(
                    dpgf_row, code, company_name, "", tco_codes_set, is_code_matched=False
                )
            )

            for idx, row in merged_df.iterrows():
                if (
                    row["row_type"] == "recap"
                    and _normalize_code(row.get("parent_code", "")) == parent_code
                ):
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
                    insertions.append((int(idx), len(insertions), new_row))
                    matched_count += 1
                    found_insertion = True
                    log.info(
                        "Insertion programmée pour article : %s dans section %s", code, parent_code
                    )
                    break

            if not found_insertion:
                # Fallback hiérarchique
                fallback_parts = parent_code.split(".")
                while len(fallback_parts) > 0 and not found_insertion:
                    fallback_parts.pop()
                    if not fallback_parts:
                        break
                    fallback_parent = ".".join(fallback_parts)
                    for idx, row in merged_df.iterrows():
                        if (
                            row["row_type"] == "recap"
                            and _normalize_code(row.get("parent_code", "")) == fallback_parent
                        ):
                            new_row = _build_new_row(
                                code,
                                dpgf_row,
                                merged_df,
                                col_u,
                                col_qu,
                                col_pu,
                                col_tot,
                                col_com,
                                fallback_parent,
                                int(idx),
                            )
                            insertions.append((int(idx), len(insertions), new_row))
                            matched_count += 1
                            found_insertion = True
                            log.info(
                                "Insertion repli : '%s' → section ancêtre '%s' (parent direct '%s' absent du modèle)",
                                code,
                                fallback_parent,
                                parent_code,
                            )
                            break

            if not found_insertion:
                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": f"Code '{code}' du DPGF non trouvé (parent inconnu)",
                    }
                )

    # P2 FIX : insertion O(n) en un seul passage au lieu de pd.concat en boucle O(n²)
    if insertions:
        ins_by_pos: dict[int, list[tuple[int, dict]]] = defaultdict(list)
        for target_idx, order, row_data in insertions:
            ins_by_pos[target_idx].append((order, row_data))
        # Trier chaque groupe par ordre original DPGF (préserve l'ordre du fichier)
        for pos in ins_by_pos:
            ins_by_pos[pos].sort(key=lambda x: x[0])

        new_rows: list[dict] = []
        for pos, record in enumerate(merged_df.to_dict("records")):
            if pos in ins_by_pos:
                for _, row_data in ins_by_pos[pos]:
                    new_rows.append(row_data)
            new_rows.append(record)

        merged_df = pd.DataFrame(new_rows).reset_index(drop=True)

    # Taux de correspondance DPGF / Template
    total_dpgf = len(dpgf_data)
    if total_dpgf > 0:
        match_rate = matched_count / total_dpgf * 100
        unmatched = total_dpgf - matched_count
        if match_rate < 50:
            msg = (
                f"DPGF ignoré — trop peu de codes correspondent au template "
                f"({matched_count}/{total_dpgf} codes matchés, soit {match_rate:.0f}%). "
                f"Vérifiez que le bon template TCO est chargé pour ce lot, "
                f"ou que les codes du DPGF entreprise suivent la même numérotation."
            )
            log.error(
                "Match rate critique %s : %.1f%% (%d/%d)",
                company_name,
                match_rate,
                matched_count,
                total_dpgf,
            )
            alerts.insert(
                0,
                {
                    "type": "error",
                    "color": "red",
                    "row": 0,
                    "code": "",
                    "message": msg,
                },
            )
            return tco_df, alerts
        elif match_rate < 90:
            msg = (
                f"Correspondance DPGF/Template partielle : {match_rate:.0f}% "
                f"— {unmatched} codes non trouvés sur {total_dpgf}"
            )
            log.warning(msg)
            alerts.insert(
                0,
                {
                    "type": "warning",
                    "color": "orange",
                    "row": 0,
                    "code": "",
                    "message": msg,
                },
            )
        log.info(
            "Match rate %s : %.1f%% (%d/%d)", company_name, match_rate, matched_count, total_dpgf
        )

    log.info("Fusion terminée : %d lignes matchées, %d non trouvées", matched_count, len(alerts))

    compute_section_totals(merged_df, col_tot, tva_rate=tva_rate)
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
            merged, comp_data["dpgf_df"], comp_name, tva_rate=tva_rate
        )
        for alert in comp_data.get("parse_alerts", []):
            alert["company"] = comp_name
        for alert in merge_alerts:
            alert["company"] = comp_name
        all_alerts.extend(comp_data.get("parse_alerts", []))
        all_alerts.extend(merge_alerts)

    return merged, all_alerts


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
    children_index: dict[str, list[int]] = {}
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if code:
            children_index.setdefault(code, []).append(idx)

    section_header_index = _build_section_index(df)

    # PERF-7 : Précalcul des sommes par préfixe parent en un seul passage O(N*depth)
    # Évite O(S×N) itérations (une par section_header) de l'ancien _get_children_total.
    parent_sums: dict[str, Decimal] = defaultdict(lambda: Decimal("0.0"))
    for code, idx_list in children_index.items():
        for idx in idx_list:
            row = df.loc[idx]
            if row["row_type"] not in ("article", "sub_section"):
                continue
            val = row.get(total_col)
            if val is None:
                continue
            try:
                v = val if isinstance(val, Decimal) else Decimal(str(val))
            except (ValueError, TypeError, ArithmeticError):  # noqa: S110
                continue
            # Propager vers tous les préfixes ancêtres (01.1.2 → 01.1 ET 01)
            parts = code.split(".")
            for i in range(1, len(parts)):
                parent_sums[".".join(parts[:i])] += v

    # Passe 1 : totaux des section_headers (lookup O(1))
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or row["row_type"] != "section_header":
            continue
        total = parent_sums.get(code, Decimal("0.0")).quantize(
            Decimal("0.01"), rounding=ROUND_HALF_UP
        )
        df.at[idx, total_col] = total

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

    # Passe 3b : vider les section_headers — le total est désormais porté
    # uniquement par la ligne recap (évite le doublon visuel section_header / recap).
    # Placé après Passe 3 car recap_summary lit encore la valeur depuis section_header.
    for idx, row in df.iterrows():
        if row["row_type"] == "section_header":
            df.at[idx, total_col] = None

    # Passe 4 : Montant HT / TVA / TTC — somme des recap_summary (= lignes du récapitulatif)
    # Les recap_summary ont été remplis en Passe 3 depuis leur section_header,
    # donc leur somme reflète exactement ce qui est affiché dans le récapitulatif.
    montant_ht = Decimal("0.0")
    for _idx, row in df.iterrows():
        if row["row_type"] == "recap_summary":
            val = row.get(total_col)
            if val is not None:
                try:
                    montant_ht += val if isinstance(val, Decimal) else Decimal(str(val))
                except (ValueError, TypeError):  # noqa: S110
                    pass

    _apply_total_lines(df, total_col, montant_ht, tva_rate)

    # Colonne de base Px_Tot_HT (si on vient de calculer une colonne entreprise)
    if total_col != "Px_Tot_HT":
        _compute_ht_tva_ttc_base(df, tva_rate)


def _compute_ht_tva_ttc_base(df: pd.DataFrame, tva_rate: float = TVA_DEFAULT) -> None:
    """Calcule Montant HT, TVA, TTC pour la colonne de base Px_Tot_HT."""
    mask = (df["row_type"] == "recap_summary") & df["Px_Tot_HT"].notna()
    montant_ht = Decimal("0.0")
    for val in df.loc[mask, "Px_Tot_HT"]:
        try:
            montant_ht += val if isinstance(val, Decimal) else Decimal(str(val))
        except (ValueError, TypeError):  # noqa: S110
            pass

    _apply_total_lines(df, "Px_Tot_HT", montant_ht, tva_rate)
