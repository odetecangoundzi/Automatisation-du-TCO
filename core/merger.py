"""
merger.py — Fusion d'un DPGF normalisé dans le TCO modèle.

Fusionne par colonne Code. Ajoute dynamiquement les colonnes entreprise.
Calcule les sous-totaux en se basant sur la hiérarchie Code (01.X → 01.X.Y...).
Les lignes "recap" (Code vide, Entete Bord_xxx_Recap) reçoivent le total
de leur section parente.
"""

from decimal import Decimal, ROUND_HALF_UP
import pandas as pd

from config import TVA_DEFAULT
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _normalize_code(code: object) -> str:
    """Normalise un code pour la comparaison."""
    return str(code).strip() if code else ""


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
    total  = Decimal("0.0")
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
                    except (ValueError, TypeError, Exception):
                        pass
    return total.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)


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
    alerts    = []

    col_qu  = f"{company_name}_Qu."
    col_pu  = f"{company_name}_Px_U_HT"
    col_tot = f"{company_name}_Px_Tot_HT"
    col_com = f"{company_name}_Commentaire"

    merged_df[col_qu]  = None
    merged_df[col_pu]  = None
    merged_df[col_tot] = None
    merged_df[col_com] = None

    if dpgf_df.empty:
        log.warning("Le DPGF de %s est vide. Aucune fusion effectuée.", company_name)
        return merged_df, [{"type": "warning", "color": "orange", "row": 0, "code": "", "message": "Fichier DPGF vide ou non reconnu."}]

    # LOT MISMATCH DETECTION
    # On regarde le premier code d'article du DPGF
    dpgf_articles = dpgf_df[dpgf_df["row_type"] == "article"]
    if not dpgf_articles.empty:
        first_code = _normalize_code(dpgf_articles.iloc[0]["Code"])
        if "." in first_code:
            dpgf_lot_prefix = first_code.split(".")[0]
            # On cherche un article dans le TCO pour comparer
            tco_articles = merged_df[merged_df["row_type"] == "article"]
            if not tco_articles.empty:
                tco_first_code = _normalize_code(tco_articles.iloc[0]["Code"])
                if "." in tco_first_code:
                    tco_lot_prefix = tco_first_code.split(".")[0]
                    if dpgf_lot_prefix != tco_lot_prefix:
                        msg = "le DPGF entreprise ne correspond pas au Template"
                        log.error("%s (Lot mismatch: DPGF=%s vs TCO=%s)", msg, dpgf_lot_prefix, tco_lot_prefix)
                        alerts.append({
                            "type": "error", "color": "red", "row": 0, "code": "", "message": msg
                        })
                        # STOP-MERGE : on retourne le TCO original (non modifié) pour cette entreprise
                        return tco_df, alerts

    # PERF-1 : index TCO codes → O(1) lookup
    tco_code_index: dict[str, int] = {}
    for idx, row in merged_df.iterrows():
        code = _normalize_code(row["Code"])
        if code and row["row_type"] not in ("empty", "recap", "recap_summary"):
            if code not in tco_code_index:
                tco_code_index[code] = idx

    # Fusion articles + sub_sections du DPGF
    dpgf_data = dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])]
    matched_count = 0
    insertions: list[tuple[int, int, dict]] = []  # (idx_to_insert_at, original_order, row_data)

    for _, dpgf_row in dpgf_data.iterrows():
        code = _normalize_code(dpgf_row["Code"])
        if not code:
            # CAS 1 FIX : alerter si une ligne sans code porte un montant non nul
            # (évite la perte silencieuse d'articles mal formatés dans le DPGF)
            px_tot = dpgf_row.get("Px_Tot_HT")
            try:
                if px_tot and Decimal(str(px_tot)) > 0:
                    desig = str(dpgf_row.get("Désignation", ""))[:60]
                    alerts.append({
                        "type": "warning",
                        "color": "orange",
                        "code": "",
                        "message": f"Ligne sans code ignorée (montant={px_tot} €) — {desig}",
                    })
            except Exception:
                pass
            continue
        if code in tco_code_index:
            idx = tco_code_index[code]
            merged_df.at[idx, col_qu]  = dpgf_row["Qu."]
            merged_df.at[idx, col_pu]  = dpgf_row["Px_U_HT"]
            merged_df.at[idx, col_tot] = dpgf_row["Px_Tot_HT"]
            merged_df.at[idx, col_com] = dpgf_row.get("Commentaire", "")
            matched_count += 1
        else:
            # DYNAMIC INSERTION
            # On tente d'insérer juste avant le récapitulatif de sa section parente
            parent_code = ".".join(code.split(".")[:-1])

            # On cherche la ligne 'recap' pour ce parent
            found_insertion = False
            for idx, row in merged_df.iterrows():
                if row["row_type"] == "recap" and _normalize_code(row.get("parent_code", "")) == parent_code:
                    new_row = {
                        "Code": code,
                        "Désignation": dpgf_row["Désignation"],
                        "Qu.": None, "U": dpgf_row.get("U", ""), "Px_U_HT": None, "Px_Tot_HT": None,
                        "Entete": dpgf_row.get("Entete", "Ouv_Art"),
                        "row_type": "article",
                        "original_row": idx,
                        "parent_code": parent_code,
                    }
                    for col in merged_df.columns:
                        if any(suffix in col for suffix in ["_Qu.", "_Px_U_HT", "_Px_Tot_HT", "_Commentaire"]):
                            new_row[col] = None

                    new_row[col_qu] = dpgf_row["Qu."]
                    new_row[col_pu] = dpgf_row["Px_U_HT"]
                    new_row[col_tot] = dpgf_row["Px_Tot_HT"]
                    new_row[col_com] = dpgf_row.get("Commentaire", "")

                    # P4 FIX : on stocke (target_idx, ordre_original, data)
                    insertions.append((idx, len(insertions), new_row))
                    matched_count += 1
                    found_insertion = True
                    log.info("Insertion programmée pour article : %s dans section %s", code, parent_code)
                    break

            if not found_insertion:
                # Fallback hiérarchique : si le parent direct est absent du modèle,
                # remonter d'un niveau à la fois jusqu'à trouver un recap ancêtre.
                # Cas typique : B2R a ajouté 03.5.2.5 / 03.5.2.6 absents du template.
                fallback_parts = parent_code.split(".")
                while len(fallback_parts) > 0 and not found_insertion:
                    fallback_parts.pop()
                    if not fallback_parts:
                        break
                    fallback_parent = ".".join(fallback_parts)
                    for idx, row in merged_df.iterrows():
                        if row["row_type"] == "recap" and _normalize_code(row.get("parent_code", "")) == fallback_parent:
                            new_row = {
                                "Code": code,
                                "Désignation": dpgf_row["Désignation"],
                                "Qu.": None, "U": dpgf_row.get("U", ""),
                                "Px_U_HT": None, "Px_Tot_HT": None,
                                "Entete": dpgf_row.get("Entete", "Ouv_Art"),
                                "row_type": "article",
                                "original_row": idx,
                                "parent_code": fallback_parent,
                            }
                            for col in merged_df.columns:
                                if any(suffix in col for suffix in ["_Qu.", "_Px_U_HT", "_Px_Tot_HT", "_Commentaire"]):
                                    new_row[col] = None
                            new_row[col_qu]  = dpgf_row["Qu."]
                            new_row[col_pu]  = dpgf_row["Px_U_HT"]
                            new_row[col_tot] = dpgf_row["Px_Tot_HT"]
                            new_row[col_com] = dpgf_row.get("Commentaire", "")
                            insertions.append((idx, len(insertions), new_row))
                            matched_count += 1
                            found_insertion = True
                            log.info(
                                "Insertion repli : '%s' → section ancêtre '%s' (parent direct '%s' absent du modèle)",
                                code, fallback_parent, parent_code,
                            )
                            break

            if not found_insertion:
                alerts.append({
                    "type":    "warning",
                    "color":   "orange",
                    "code":    code,
                    "message": f"Code '{code}' du DPGF non trouvé (parent inconnu)",
                })

    # Application des insertions en ordre décroissant pour préserver les positions
    # P4 FIX : tri par (target_idx DESC, ordre_original DESC) — pour le même target_idx,
    # on insère les articles en ordre inverse afin qu'après toutes les insertions
    # leur ordre final corresponde à l'ordre original du DPGF.
    if insertions:
        insertions.sort(key=lambda x: (x[0], x[1]), reverse=True)
        for target_idx, _order, row_data in insertions:
            part1 = merged_df.iloc[:target_idx]
            part2 = merged_df.iloc[target_idx:]
            merged_df = pd.concat([part1, pd.DataFrame([row_data]), part2], ignore_index=True)

    # ------------------------------------------------------------------
    # Point 4 : Taux de correspondance DPGF / Template
    # Seuil critique < 50 % → erreur bloquante (avertissement fort).
    # Seuil partiel 50-90 % → warning.
    # ≥ 90 % → OK, pas d'alerte de correspondance.
    # ------------------------------------------------------------------
    total_dpgf = len(dpgf_data)
    if total_dpgf > 0:
        match_rate = matched_count / total_dpgf * 100
        unmatched  = total_dpgf - matched_count
        if match_rate < 50:
            msg = "le DPGF entreprise ne correspond pas au Template"
            log.error("%s (Match rate critique: %.1f%%)", msg, match_rate)
            alerts.insert(0, {
                "type": "error", "color": "red", "row": 0, "code": "",
                "message": msg,
            })
            # STOP-MERGE : trop peu de correspondances pour être fiable
            return tco_df, alerts
        elif match_rate < 90:
            msg = (
                f"Correspondance DPGF/Template partielle : {match_rate:.0f}% "
                f"— {unmatched} codes non trouvés sur {total_dpgf}"
            )
            log.warning(msg)
            alerts.insert(0, {
                "type": "warning", "color": "orange", "row": 0, "code": "",
                "message": msg,
            })
        log.info("Match rate %s : %.1f%% (%d/%d)", company_name, match_rate, matched_count, total_dpgf)

    log.info(
        "Fusion terminée : %d lignes matchées, %d non trouvées",
        matched_count, len(alerts)
    )

    _compute_section_totals(merged_df, col_tot, tva_rate=tva_rate)
    return merged_df, alerts


# ---------------------------------------------------------------------------
# Calcul des sous-totaux
# ---------------------------------------------------------------------------

def _compute_section_totals(
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

    # Passe 1 : calculer les totaux des section_headers
    for idx, row in df.iterrows():
        code = _normalize_code(row["Code"])
        if not code or row["row_type"] != "section_header":
            continue
        total = _get_children_total(df, code, total_col, children_index)
        df.at[idx, total_col] = total

    # Passe 2 : propager vers les lignes recap
    for idx, row in df.iterrows():
        if row["row_type"] != "recap":
            continue
        parent = _normalize_code(row.get("parent_code", ""))
        if parent and parent in section_header_index:
            s_idx = section_header_index[parent]
            val   = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val

    # Passe 3 : recap_summary
    for idx, row in df.iterrows():
        if row["row_type"] != "recap_summary":
            continue
        code = _normalize_code(row["Code"])
        if code and code in section_header_index:
            s_idx = section_header_index[code]
            val   = df.at[s_idx, total_col]
            if val is not None:
                df.at[idx, total_col] = val

    # Passe 4 : Montant HT / TVA / TTC
    montant_ht = Decimal("0.0")
    for idx, row in df.iterrows():
        if row["row_type"] == "section_header":
            val = row.get(total_col)
            if val is not None:
                try:
                    if isinstance(val, Decimal):
                        montant_ht += val
                    else:
                        montant_ht += Decimal(str(val))
                except (ValueError, TypeError, Exception):
                    pass

    if montant_ht > 0:
        d_tva_rate  = Decimal(str(tva_rate))
        tva         = (montant_ht * d_tva_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
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
    else:
        # Si montant_ht est 0 ou non calculable, on met 0 par défaut pour les lignes de total
        for idx, row in df.iterrows():
            if row["row_type"] == "total_line":
                df.at[idx, total_col] = 0.0

    # Colonne de base Px_Tot_HT (si on vient de calculer une colonne entreprise)
    if total_col != "Px_Tot_HT":
        _compute_ht_tva_ttc_base(df, tva_rate)


def _compute_ht_tva_ttc_base(df: pd.DataFrame, tva_rate: float = TVA_DEFAULT) -> None:
    """Calcule Montant HT, TVA, TTC pour la colonne de base Px_Tot_HT."""
    mask = (
        (df["row_type"] == "section_header")
        & df["Px_Tot_HT"].notna()
    )
    
    montant_ht = Decimal("0.0")
    for val in df.loc[mask, "Px_Tot_HT"]:
        try:
            if isinstance(val, Decimal):
                montant_ht += val
            else:
                montant_ht += Decimal(str(val))
        except (ValueError, TypeError, Exception):
            pass

    if montant_ht <= 0:
        return

    d_tva_rate  = Decimal(str(tva_rate))
    tva         = (montant_ht * d_tva_rate).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    montant_ttc = (montant_ht + tva).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    term_map    = {"montant ht": montant_ht, "tva": tva, "ttc": montant_ttc}

    for idx, row in df.iterrows():
        if row["row_type"] != "total_line":
            continue
        desig = str(row.get("Désignation", "")).strip().lower()
        for key, val in term_map.items():
            if key in desig:
                df.at[idx, "Px_Tot_HT"] = val
                break
