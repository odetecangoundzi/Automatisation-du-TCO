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
        "is_extra_line": True, # Tag to identify lines not in original model
    }
    for col in merged_df.columns:
        if any(suffix in col for suffix in ["_Qu.", "_Px_U_HT", "_Px_Tot_HT", "_Commentaire"]):
            new_row[col] = None
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

    col_qu = f"{company_name}_Qu."
    col_pu = f"{company_name}_Px_U_HT"
    col_tot = f"{company_name}_Px_Tot_HT"
    col_com = f"{company_name}_Commentaire"

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

    # Fusion articles + sub_sections du DPGF
    dpgf_data = dpgf_df[dpgf_df["row_type"].isin(["article", "sub_section"])]
    matched_count = 0
    insertions: list[tuple[int, int, dict]] = []  # (target_idx, order, row_data)

    for _, dpgf_row in dpgf_data.iterrows():
        code = _normalize_code(dpgf_row["Code"])
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

                # Check Quantity mismatch (ignoring 0 vs None handling)
                try:
                    tco_qu_val = float(tco_qu) if tco_qu is not None else 0.0
                    dpgf_qu_val = float(dpgf_qu) if dpgf_qu is not None else 0.0
                    if tco_qu_val != dpgf_qu_val and tco_qu_val > 0:
                        alerts.append({
                            "type": "warning",
                            "color": "orange",
                            "code": code,
                            "message": f"Quantité divergente par rapport au modèle ({tco_qu} vs {dpgf_qu})",
                        })
                except (ValueError, TypeError):
                    pass

                # Check Unit mismatch
                if tco_u and dpgf_u and tco_u != dpgf_u:
                    alerts.append({
                        "type": "warning",
                        "color": "orange",
                        "code": code,
                        "message": f"Unité modifiée par rapport au modèle ({tco_u} vs {dpgf_u})",
                    })

            matched_count += 1
        else:
            parent_code = ".".join(code.split(".")[:-1])
            found_insertion = False

            for idx, row in merged_df.iterrows():
                if (
                    row["row_type"] == "recap"
                    and _normalize_code(row.get("parent_code", "")) == parent_code
                ):
                    new_row = _build_new_row(
                        code,
                        dpgf_row,
                        merged_df,
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
            log.error("Match rate critique %s : %.1f%% (%d/%d)", company_name, match_rate, matched_count, total_dpgf)
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
