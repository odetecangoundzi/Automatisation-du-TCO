"""
parser_dpgf.py — Normalisation et extraction des annotations des DPGF entreprise.

Gère les textes dans les cellules numériques (SANS OBJET, COMPRIS, nc, P-M),
extrait les annotations en colonne Commentaire, et détecte les erreurs.
Utilise la colonne Entete (col M) pour classifier chaque ligne.
"""

from __future__ import annotations

import re
from decimal import ROUND_HALF_UP, Decimal

import pandas as pd

from config import TOTAL_TOLERANCE_ABS, TOTAL_TOLERANCE_REL
from core.utils import classify_row, find_column_index, open_excel_file
from logger import get_logger

log = get_logger(__name__)


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------

KEYWORDS = {
    "sans objet": "so",
    "compris": "compris",
    "nc": "nc",
    "p-m": "pm",
    "inclus": "inclus",
    "néant": "néant",
    "so": "so",
    "pm": "pm",
}


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------


def _looks_numeric(val: object) -> bool:
    """Retourne True si la valeur représente un nombre réel (int, float ou str numérique).
    Utilisé pour detect has_price même quand dtype=object préserve les strings.
    """
    if pd.isna(val):
        return False
    if isinstance(val, (int, float)):
        return True
    s = str(val).strip()
    if not s:
        return False
    # Mots-clés textuels (SANS OBJET, P-M...) ne sont pas des prix
    s_lower = s.lower()
    if any(kw in s_lower for kw in KEYWORDS):
        return False
    cleaned = s.replace(" ", "").replace(" ", "").replace(",", ".")
    try:
        float(cleaned)
        return True
    except ValueError:
        return False


def _clean_numeric(value: int | float | str | None) -> tuple[Decimal, str]:
    """
    Nettoie une valeur potentiellement numérique.
    Retourne (nombre_decimal, texte_annotation).
    """
    if pd.isna(value):
        return Decimal("0.0"), ""
    if isinstance(value, (int, float)):
        return Decimal(str(value)), ""

    text = str(value).strip()
    if not text:
        return Decimal("0.0"), ""

    # Limite de longueur pour éviter ReDoS et surcharge mémoire
    if len(text) > 500:
        log.warning("Valeur trop longue (%d chars) tronquée à 500", len(text))
        text = text[:500]

    text_lower = text.lower().strip()
    for keyword, abbrev in KEYWORDS.items():
        if keyword in text_lower:
            return Decimal("0.0"), abbrev

    cleaned = text.replace(" ", "").replace("\u00a0", "").replace(",", ".")

    # On cherche tous les nombres
    matches = re.findall(r"-?\d+(?:\.\d+)?", cleaned)
    if matches:
        try:
            # P2 FIX: On prend le DERNIER nombre (souvent le prix ou la quantité finale)
            # car les premiers nombres sont souvent des indices de lot ou d'article.
            number_str = matches[-1]
            number = Decimal(number_str)

            # Reconstruction du commentaire en excluant la DERNIÈRE occurrence de ce nombre
            # (cohérent avec matches[-1] — on retire la dernière, pas la première)
            last_pos = cleaned.rfind(number_str)
            remaining = (cleaned[:last_pos] + cleaned[last_pos + len(number_str) :]).strip()
            remaining = remaining.strip("()[]{}/ ")
            # On nettoie un peu le résidu s'il reste des points ou tirets
            remaining = re.sub(r"^[.\-:]+", "", remaining).strip()

            return number, remaining
        except (ValueError, ArithmeticError):  # noqa: S110
            pass

    return Decimal("0.0"), text


def _check_total_coherence(
    qu_val: Decimal, pu_val: Decimal, total_val: Decimal, row_idx: int, code: str
) -> dict | None:
    """Vérifie que Qu × PU ≈ Total (tolérance absolue ET relative)."""
    if qu_val and pu_val and total_val:
        try:
            expected = (qu_val * pu_val).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            actual = total_val.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
            if actual != 0:
                abs_diff = abs(expected - actual)
                rel_diff = abs_diff / abs(actual)
                if abs_diff > Decimal(str(TOTAL_TOLERANCE_ABS)) and rel_diff > Decimal(
                    str(TOTAL_TOLERANCE_REL)
                ):
                    log.warning(
                        "Total incohérent ligne %d code=%s : %s × %s = %s ≠ %s",
                        row_idx,
                        code,
                        qu_val,
                        pu_val,
                        expected,
                        actual,
                    )
                    return {
                        "type": "error",
                        "color": "red",
                        "row": row_idx,
                        "code": code,
                        "message": (
                            f"Total incohérent : {qu_val} × {pu_val} = "
                            f"{expected} ≠ {actual} "
                            f"(écart {abs_diff} €)"
                        ),
                    }
        except (ValueError, TypeError):  # noqa: S110
            pass
    return None


# ---------------------------------------------------------------------------
# Main parser
# ---------------------------------------------------------------------------


def parse_dpgf(filepath: str) -> tuple[pd.DataFrame, list[dict]]:
    """
    Lit et normalise un fichier DPGF entreprise (XLSX, XLS, XLSB).

    Returns:
        dpgf_df (DataFrame) : DataFrame normalisé
        alerts  (list)      : liste d'alertes
    """
    log.info("Lecture DPGF : %s", filepath)

    try:
        # open_excel_file : détecte engine, feuille et en-tête en un seul appel
        # (2 lectures au lieu de 3 — le probe est réutilisé comme df_raw)
        xl_file, sheet_name, _df_raw, header_row_idx, _engine_kwargs = open_excel_file(filepath)

        # Lecture finale avec skiprows=header_row_idx (1 seule lecture supplémentaire)
        df_data = xl_file.parse(
            sheet_name,
            skiprows=header_row_idx,
            dtype=object,  # preserve codes comme strings
        )
    except Exception as e:
        log.error("Erreur de structure DPGF: %s", e)
        return pd.DataFrame(), [
            {"type": "error", "color": "red", "row": 0, "code": "", "message": str(e)}
        ]

    alerts = []
    rows = []
    current_section_code = ""

    # Mapping dynamique des colonnes — mots-clés étendus pour DPGFs hétérogènes
    idx_code = find_column_index(df_data, ["code", "n°", "n°.", "n° de prix", "num", "indice"], 0)
    idx_desig = find_column_index(df_data, ["désignation", "designation", "libellé", "libelle"], 1)
    idx_qu = find_column_index(df_data, ["qu.", "quantité", "qte", "qté", "qt", "quantite", "q"], 2)
    idx_u = find_column_index(df_data, ["u", "unité", "unite"], 3)
    idx_pu = find_column_index(
        df_data, ["px u", "p.u", "prix u", "prix unitaire", "pu", "px u."], 4
    )
    idx_tot = find_column_index(
        df_data,
        ["px tot", "total ht", "prix tot", "montant ht", "total h", "pt", "px tot.", "prix tot."],
        5,
    )
    idx_entete = find_column_index(df_data, ["entete", "entête"])  # COL_NOT_FOUND (-1) si absent

    for idx_in_df, xl_row in df_data.iterrows():
        row_idx = idx_in_df + header_row_idx + 2  # conversion en 1-indexed Excel row

        if len(xl_row) <= max(idx_code, idx_desig, idx_qu, idx_pu, idx_tot):
            continue

        code_raw = xl_row.iloc[idx_code]
        desig_raw = xl_row.iloc[idx_desig]
        cc_raw = xl_row.iloc[idx_qu]
        u = xl_row.iloc[idx_u]
        px_u_raw = xl_row.iloc[idx_pu]
        px_tot_raw = xl_row.iloc[idx_tot]
        entete = xl_row.iloc[idx_entete] if (idx_entete >= 0 and len(xl_row) > idx_entete) else None

        code_str = str(code_raw).strip() if pd.notna(code_raw) else ""
        desig_str = str(desig_raw).strip() if pd.notna(desig_raw) else ""
        ent_str = str(entete).strip() if pd.notna(entete) else ""

        # has_price = True si les deux valeurs sont numériques (int, float ou str numérique)
        # _looks_numeric gère correctement dtype=object (strings) et les mots-clés textuels
        has_price = _looks_numeric(cc_raw) and _looks_numeric(px_u_raw)

        row_type = classify_row(code_str, desig_str, ent_str, has_price=has_price)

        if row_type == "section_header":
            current_section_code = code_str

        parent_code = current_section_code if row_type == "recap" else ""

        if row_type in ("article", "sub_section"):
            qu_val, qu_comment = _clean_numeric(cc_raw)
            pu_val, pu_comment = _clean_numeric(px_u_raw)
            tot_val, tot_comment = _clean_numeric(px_tot_raw)
        else:
            qu_val = (
                Decimal(str(cc_raw))
                if isinstance(cc_raw, (int, float)) and not pd.isna(cc_raw)
                else Decimal("0.0")
            )
            pu_val = (
                Decimal(str(px_u_raw))
                if isinstance(px_u_raw, (int, float)) and not pd.isna(px_u_raw)
                else Decimal("0.0")
            )
            tot_val = (
                Decimal(str(px_tot_raw))
                if isinstance(px_tot_raw, (int, float)) and not pd.isna(px_tot_raw)
                else Decimal("0.0")
            )
            qu_comment = pu_comment = tot_comment = ""

        comments = [c for c in [qu_comment, pu_comment, tot_comment] if c]
        commentaire = "; ".join(comments) if comments else ""

        if row_type == "article" and code_str:
            if qu_comment or pu_comment or tot_comment:
                kw_found = any(
                    c.lower() in KEYWORDS or c.lower() in KEYWORDS.values()
                    for c in [qu_comment, pu_comment, tot_comment]
                    if c
                )
                alert_type = ("info", "blue") if kw_found else ("warning", "yellow")
                msg = (
                    f"Mot-clé détecté : {commentaire}"
                    if kw_found
                    else f"Texte dans champ numérique : {commentaire}"
                )
                alerts.append(
                    {
                        "type": alert_type[0],
                        "color": alert_type[1],
                        "row": row_idx,
                        "code": code_str,
                        "message": msg,
                    }
                )
                log.debug("Alerte %s code=%s : %s", alert_type[0], code_str, msg)

            alert = _check_total_coherence(qu_val, pu_val, tot_val, row_idx, code_str)
            if alert:
                alerts.append(alert)

            # Point 7 : unité manquante (article avec montant mais sans unité)
            u_str = str(u).strip() if pd.notna(u) else ""
            if not u_str and tot_val > 0:
                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "row": row_idx,
                        "code": code_str,
                        "message": f"Unité manquante (Px_Tot={tot_val} €)",
                    }
                )

        rows.append(
            {
                "Code": code_str,
                "Désignation": desig_str,
                "Qu.": qu_val,
                "U": str(u).strip() if pd.notna(u) else "",
                "Px_U_HT": pu_val,
                "Px_Tot_HT": tot_val,
                "Commentaire": commentaire,
                "Entete": ent_str,
                "row_type": row_type,
                "original_row": row_idx,
                "parent_code": parent_code,
            }
        )

    dpgf_df = pd.DataFrame(rows)

    # ------------------------------------------------------------------
    # Point 5 : Détection et gestion des codes dupliqués dans le DPGF
    # Option A : suffixe technique _DUPxx pour unicité, Code_source conserve l'original.
    # La 1ère occurrence garde le code intact (matchera le TCO).
    # Les suivantes reçoivent un suffixe et seront insérées comme nouveaux articles.
    # ------------------------------------------------------------------
    if not dpgf_df.empty:
        art_mask = dpgf_df["row_type"].isin(["article", "sub_section"])
        art_sub = dpgf_df[art_mask & (dpgf_df["Code"] != "")]
        dup_codes = set(art_sub[art_sub.duplicated(subset=["Code"], keep=False)]["Code"].unique())
        if dup_codes:
            # Ajouter la colonne Code_source (original) avant de modifier Code
            dpgf_df.insert(
                dpgf_df.columns.get_loc("Code") + 1,
                "Code_source",
                dpgf_df["Code"].copy(),
            )
            code_seen: dict[str, int] = {}
            for idx in dpgf_df.index:
                c = dpgf_df.at[idx, "Code"]
                if c in dup_codes and dpgf_df.at[idx, "row_type"] in ("article", "sub_section"):
                    code_seen[c] = code_seen.get(c, 0) + 1
                    if code_seen[c] > 1:
                        dpgf_df.at[idx, "Code"] = f"{c}_DUP{code_seen[c]:02d}"
            # Générer un warning par code dupliqué
            for dup_code in sorted(dup_codes):
                dup_rows = dpgf_df[dpgf_df["Code_source"] == dup_code]
                n = len(dup_rows)
                desigs = " | ".join(dup_rows["Désignation"].astype(str).str[:30])
                alerts.append(
                    {
                        "type": "warning",
                        "color": "orange",
                        "row": int(dup_rows.iloc[0].get("original_row", 0)),
                        "code": dup_code,
                        "message": f"Code dupliqué ({n}×) — {desigs}",
                    }
                )
                log.warning("Code dupliqué %s (%d occurrences)", dup_code, n)

    log.info(
        "DPGF parsé : %d lignes, %d alertes",
        len(dpgf_df),
        len(alerts),
    )
    return dpgf_df, alerts
