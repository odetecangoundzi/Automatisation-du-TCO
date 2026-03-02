"""
test_merger.py — Tests de la logique de fusion DPGF→TCO.

Couvre : merge_company_into_tco, merge_all_companies,
         compute_section_totals, _normalize_code.
"""

from decimal import Decimal

import pandas as pd
import pytest

from core.merger import (
    _normalize_code,
    compute_section_totals,
    merge_all_companies,
    merge_company_into_tco,
)

# ---------------------------------------------------------------------------
# Fixtures locales
# ---------------------------------------------------------------------------


@pytest.fixture
def tco_two_articles(minimal_tco_df):
    """TCO de la fixture conftest (section '1' + article '1.1' + recap)."""
    return minimal_tco_df


def _make_tco_no_recap(article_codes: list[str]) -> pd.DataFrame:
    """TCO sans ligne recap — l'insertion hiérarchique est impossible.

    Sans recap, les codes DPGF non trouvés ne peuvent pas être insérés,
    ce qui permet de tester des taux de match réels < 100%.
    """
    rows = [
        {
            "Code": "1",
            "Désignation": "Section A",
            "Qu.": Decimal("0"),
            "U": "",
            "Px_U_HT": Decimal("0"),
            "Px_Tot_HT": Decimal("0"),
            "Entete": "Bd_01_Bord",
            "row_type": "section_header",
            "original_row": 1,
            "parent_code": "",
        }
    ]
    for i, code in enumerate(article_codes, start=2):
        rows.append(
            {
                "Code": code,
                "Désignation": f"Article {code}",
                "Qu.": Decimal("10"),
                "U": "u",
                "Px_U_HT": Decimal("10"),
                "Px_Tot_HT": Decimal("100"),
                "Entete": "Ouv_01_Art",
                "row_type": "article",
                "original_row": i,
                "parent_code": "",
            }
        )
    # Pas de ligne recap → aucun point d'insertion hiérarchique
    return pd.DataFrame(rows)


def _make_dpgf_articles(codes: list[str]) -> pd.DataFrame:
    """DPGF minimal avec les codes donnés."""
    rows = []
    for i, code in enumerate(codes, start=1):
        rows.append(
            {
                "Code": code,
                "Désignation": f"Article {code}",
                "Qu.": Decimal("10"),
                "U": "u",
                "Px_U_HT": Decimal("10"),
                "Px_Tot_HT": Decimal("100"),
                "Commentaire": "",
                "Entete": "Ouv_01_Art",
                "row_type": "article",
                "original_row": i,
                "parent_code": "",
            }
        )
    return pd.DataFrame(rows)


@pytest.fixture
def tco_lot2():
    """TCO avec articles de lot 2 (codes 2.x) pour test lot mismatch."""
    return pd.DataFrame(
        [
            {
                "Code": "2",
                "Désignation": "Section B",
                "Qu.": Decimal("0"),
                "U": "",
                "Px_U_HT": Decimal("0"),
                "Px_Tot_HT": Decimal("0"),
                "Entete": "Bd_02_Bord",
                "row_type": "section_header",
                "original_row": 2,
                "parent_code": "",
            },
            {
                "Code": "2.1",
                "Désignation": "Article lot 2",
                "Qu.": Decimal("5"),
                "U": "u",
                "Px_U_HT": Decimal("50"),
                "Px_Tot_HT": Decimal("250"),
                "Entete": "Ouv_02_Art",
                "row_type": "article",
                "original_row": 3,
                "parent_code": "",
            },
            {
                "Code": "",
                "Désignation": "Total section B",
                "Qu.": Decimal("0"),
                "U": "",
                "Px_U_HT": Decimal("0"),
                "Px_Tot_HT": Decimal("0"),
                "Entete": "Bord_02_Recap",
                "row_type": "recap",
                "original_row": 4,
                "parent_code": "2",
            },
        ]
    )


@pytest.fixture
def dpgf_lot1(minimal_dpgf_df):
    """DPGF avec 1 article du lot 1."""
    return minimal_dpgf_df


@pytest.fixture
def dpgf_lot2():
    """DPGF avec article du lot 2 (pour déclencher lot mismatch vs lot 1)."""
    return pd.DataFrame(
        [
            {
                "Code": "2.1",
                "Désignation": "Article lot 2 DPGF",
                "Qu.": Decimal("5"),
                "U": "u",
                "Px_U_HT": Decimal("60"),
                "Px_Tot_HT": Decimal("300"),
                "Commentaire": "",
                "Entete": "Ouv_02_Art",
                "row_type": "article",
                "original_row": 2,
                "parent_code": "",
            }
        ]
    )


# ---------------------------------------------------------------------------
# Tests comportement de base
# ---------------------------------------------------------------------------


class TestMergeBasic:
    def test_merge_adds_company_columns(self, minimal_tco_df, minimal_dpgf_df):
        """4 colonnes {company}_* ajoutées après fusion."""
        merged, _ = merge_company_into_tco(minimal_tco_df, minimal_dpgf_df, "ACME")
        for col in ["ACME_Qu.", "ACME_Px_U_HT", "ACME_Px_Tot_HT", "ACME_Commentaire"]:
            assert col in merged.columns, f"Colonne {col} manquante"

    def test_merge_prices_correctly_matched(self, minimal_tco_df, minimal_dpgf_df):
        """Le prix de l'article '1.1' est correctement transféré."""
        merged, _ = merge_company_into_tco(minimal_tco_df, minimal_dpgf_df, "ACME")
        row = merged[merged["Code"] == "1.1"].iloc[0]
        assert row["ACME_Px_Tot_HT"] == Decimal("250")
        assert row["ACME_Px_U_HT"] == Decimal("25")

    def test_merge_section_total_propagated(self, minimal_tco_df, minimal_dpgf_df):
        """La section_header '1' est vidée (doublon éliminé), le recap porte le total."""
        merged, _ = merge_company_into_tco(minimal_tco_df, minimal_dpgf_df, "ACME")
        section_row = merged[merged["Code"] == "1"].iloc[0]
        # Passe 2b : section_header vidée pour éviter le doublon visuel
        assert section_row["ACME_Px_Tot_HT"] is None
        # Le total est porté par la ligne recap
        recap_row = merged[merged["row_type"] == "recap"].iloc[0]
        assert recap_row["ACME_Px_Tot_HT"] == Decimal("250")


# ---------------------------------------------------------------------------
# Tests DPGF vide
# ---------------------------------------------------------------------------


class TestMergeEmptyDpgf:
    def test_empty_dpgf_returns_tco_with_columns(self, minimal_tco_df):
        """DPGF vide → TCO retourné avec colonnes entreprise + alerte 'vide'."""
        merged, alerts = merge_company_into_tco(minimal_tco_df, pd.DataFrame(), "EMPTY_CO")
        assert not merged.empty
        assert "EMPTY_CO_Px_Tot_HT" in merged.columns
        assert len(alerts) == 1
        assert "vide" in alerts[0]["message"].lower()


# ---------------------------------------------------------------------------
# Tests lot mismatch
# ---------------------------------------------------------------------------


class TestLotMismatch:
    def test_lot_mismatch_returns_original_tco(self, minimal_tco_df, dpgf_lot2):
        """DPGF lot 2 vs TCO lot 1 → TCO original retourné, alerte error."""
        # minimal_tco_df a article "1.1" (lot 1), dpgf_lot2 a "2.1" (lot 2)
        merged, alerts = merge_company_into_tco(minimal_tco_df, dpgf_lot2, "COMP_LOT2")
        # Retour du TCO original → pas de colonnes entreprise
        assert "COMP_LOT2_Px_Tot_HT" not in merged.columns
        error_alerts = [a for a in alerts if a["type"] == "error"]
        assert len(error_alerts) >= 1
        assert "DPGF ignoré" in error_alerts[0]["message"]


# ---------------------------------------------------------------------------
# Tests match rate
# ---------------------------------------------------------------------------


class TestMatchRate:
    def test_match_rate_below_50_returns_original(self):
        """TCO 1 article '1.1', DPGF 5 articles → 1/5 = 20% → TCO original, alerte error.

        Sans ligne recap dans le TCO, l'insertion hiérarchique est impossible :
        les 4 codes DPGF non trouvés restent non matchés → taux réel de 20%.
        """
        tco = _make_tco_no_recap(["1.1"])
        dpgf = _make_dpgf_articles(["1.1", "1.2", "1.3", "1.4", "1.5"])
        merged, alerts = merge_company_into_tco(tco, dpgf, "COMP_LOW")
        # Match rate 20% < 50% → retour TCO original sans colonnes entreprise
        assert "COMP_LOW_Px_Tot_HT" not in merged.columns
        error_alerts = [a for a in alerts if a["type"] == "error"]
        assert len(error_alerts) >= 1

    def test_match_rate_50_to_90_generates_warning(self):
        """TCO 4 articles, DPGF 5 articles → 4/5 = 80% → alerte warning avec %.

        Sans ligne recap, le 5e code DPGF ne peut pas être inséré
        → taux = 80%, entre 50% et 90% → warning généré.
        """
        tco = _make_tco_no_recap(["1.1", "1.2", "1.3", "1.4"])
        dpgf = _make_dpgf_articles(["1.1", "1.2", "1.3", "1.4", "1.5"])
        merged, alerts = merge_company_into_tco(tco, dpgf, "COMP_MED")
        warn_alerts = [a for a in alerts if a["type"] == "warning" and "%" in a.get("message", "")]
        assert len(warn_alerts) >= 1

    def test_match_rate_100_no_match_warning(self, minimal_tco_df, minimal_dpgf_df):
        """100% match → aucune alerte de match rate."""
        _, alerts = merge_company_into_tco(minimal_tco_df, minimal_dpgf_df, "COMP_FULL")
        match_alerts = [a for a in alerts if "%" in a.get("message", "")]
        assert len(match_alerts) == 0


# ---------------------------------------------------------------------------
# Tests insertion hiérarchique
# ---------------------------------------------------------------------------


class TestHierarchicalInsertion:
    def test_new_code_inserted_before_recap(self, minimal_tco_df):
        """Article '1.3' absent du TCO → inséré avant la ligne recap de la section '1'."""
        dpgf_new = pd.DataFrame(
            [
                {
                    "Code": "1.3",
                    "Désignation": "Nouvel article",
                    "Qu.": Decimal("5"),
                    "U": "m²",
                    "Px_U_HT": Decimal("100"),
                    "Px_Tot_HT": Decimal("500"),
                    "Commentaire": "",
                    "Entete": "",
                    "row_type": "article",
                    "original_row": 2,
                    "parent_code": "",
                }
            ]
        )
        merged, _ = merge_company_into_tco(minimal_tco_df, dpgf_new, "NEW_CO")
        assert "1.3" in merged["Code"].values
        # Vérifier que 1.3 est positionné avant le recap de la section 1
        idx_new = merged[merged["Code"] == "1.3"].index[0]
        idx_recap = merged[merged["row_type"] == "recap"].index[0]
        assert idx_new < idx_recap

    def test_normalize_code_leading_zeros_match(self, minimal_tco_df):
        """TCO '1.1' vs DPGF avec zéros de tête '01.1' → ils matchent."""
        # minimal_tco_df a code "1.1"
        dpgf_padded = pd.DataFrame(
            [
                {
                    "Code": "01.1",  # zéro de tête
                    "Désignation": "Article normalisé",
                    "Qu.": Decimal("10"),
                    "U": "u",
                    "Px_U_HT": Decimal("30"),
                    "Px_Tot_HT": Decimal("300"),
                    "Commentaire": "",
                    "Entete": "",
                    "row_type": "article",
                    "original_row": 2,
                    "parent_code": "",
                }
            ]
        )
        merged, alerts = merge_company_into_tco(minimal_tco_df, dpgf_padded, "NORM_CO")
        row = merged[merged["Code"] == "1.1"]
        assert not row.empty
        # Le prix doit avoir été transféré
        assert row.iloc[0]["NORM_CO_Px_Tot_HT"] == Decimal("300")


# ---------------------------------------------------------------------------
# Tests _normalize_code directement
# ---------------------------------------------------------------------------


class TestNormalizeCode:
    def test_leading_zeros_stripped(self):
        assert _normalize_code("01") == "1"
        assert _normalize_code("01.10") == "1.10"
        assert _normalize_code("03.5.2") == "3.5.2"

    def test_float_integer_excel(self):
        """Float Excel entier : 1.0 → '1'."""
        assert _normalize_code(1.0) == "1"

    def test_float_decimal_excel(self):
        """Float Excel décimal : 1.1 → '1.1'."""
        assert _normalize_code(1.1) == "1.1"

    def test_none_returns_empty(self):
        assert _normalize_code(None) == ""

    def test_nan_returns_empty(self):

        assert _normalize_code(float("nan")) == ""


# ---------------------------------------------------------------------------
# Tests compute_section_totals
# ---------------------------------------------------------------------------


class TestComputeSectionTotals:
    def test_four_passes_correct(self):
        """Vérifie les 4 passes de recalcul après fusion."""
        from decimal import Decimal as D

        df = pd.DataFrame(
            [
                {
                    "Code": "1",
                    "Désignation": "Section A",
                    "row_type": "section_header",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "Bd_01_Bord",
                    "original_row": 2,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": None,
                },
                {
                    "Code": "1.1",
                    "Désignation": "Article 1",
                    "row_type": "article",
                    "Qu.": D("10"),
                    "Px_U_HT": D("20"),
                    "Px_Tot_HT": D("200"),
                    "U": "u",
                    "Entete": "Ouv_01_Art",
                    "original_row": 3,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": D("300"),
                },
                {
                    "Code": "",
                    "Désignation": "Total section A",
                    "row_type": "recap",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "Bord_01_Recap",
                    "original_row": 4,
                    "parent_code": "1",
                    "COMP_Px_Tot_HT": None,
                },
                {
                    "Code": "1",
                    "Désignation": "Recap Summary Section A",
                    "row_type": "recap_summary",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "Bord_01_Recap",
                    "original_row": 4,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": None,
                },
                {
                    "Code": "",
                    "Désignation": "Montant HT",
                    "row_type": "total_line",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "LignesTot_01",
                    "original_row": 5,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": None,
                },
                {
                    "Code": "",
                    "Désignation": "TVA 20%",
                    "row_type": "total_line",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "LignesTot_02",
                    "original_row": 6,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": None,
                },
                {
                    "Code": "",
                    "Désignation": "Montant TTC",
                    "row_type": "total_line",
                    "Qu.": D("0"),
                    "Px_U_HT": D("0"),
                    "Px_Tot_HT": D("0"),
                    "U": "",
                    "Entete": "LignesTot_03",
                    "original_row": 7,
                    "parent_code": "",
                    "COMP_Px_Tot_HT": None,
                },
            ]
        )

        compute_section_totals(df, "COMP_Px_Tot_HT", tva_rate=0.20)

        # Passe 1→2b : section_header vidée, recap porte le total
        section_total = df[(df["Code"] == "1") & (df["row_type"] == "section_header")][
            "COMP_Px_Tot_HT"
        ].iloc[0]
        assert section_total is None
        recap_total = df[df["row_type"] == "recap"]["COMP_Px_Tot_HT"].iloc[0]
        assert recap_total == Decimal("300")

        # Passe 4 : Montant HT = 300, TVA = 60, TTC = 360
        ht_row = df[df["Désignation"] == "Montant HT"]["COMP_Px_Tot_HT"].iloc[0]
        tva_row = df[df["Désignation"] == "TVA 20%"]["COMP_Px_Tot_HT"].iloc[0]
        ttc_row = df[df["Désignation"] == "Montant TTC"]["COMP_Px_Tot_HT"].iloc[0]

        assert ht_row == Decimal("300")
        assert tva_row == Decimal("60.00")
        assert ttc_row == Decimal("360.00")


# ---------------------------------------------------------------------------
# Tests merge_all_companies
# ---------------------------------------------------------------------------


class TestMergeAllCompanies:
    def test_two_companies_eight_columns(self, minimal_tco_df, minimal_dpgf_df):
        """Fusion 2 entreprises → 8 colonnes entreprise (4 × 2)."""
        dpgf2 = minimal_dpgf_df.copy()
        dpgf2["Px_U_HT"] = Decimal("30")
        dpgf2["Px_Tot_HT"] = Decimal("300")

        company_data = {
            "COMP1": {"dpgf_df": minimal_dpgf_df, "parse_alerts": [], "filename": "f1.xlsx"},
            "COMP2": {"dpgf_df": dpgf2, "parse_alerts": [], "filename": "f2.xlsx"},
        }
        merged, all_alerts = merge_all_companies(minimal_tco_df, company_data, tva_rate=0.20)

        for col in ["COMP1_Px_Tot_HT", "COMP2_Px_Tot_HT", "COMP1_Qu.", "COMP2_Qu."]:
            assert col in merged.columns

    def test_alerts_tagged_with_company_name(self, minimal_tco_df):
        """Les alertes sont taguées avec le nom de l'entreprise."""
        company_data = {
            "TAGGED_CO": {
                "dpgf_df": pd.DataFrame(),
                "parse_alerts": [{"type": "info", "message": "test alert"}],
                "filename": "f.xlsx",
            }
        }
        _, all_alerts = merge_all_companies(minimal_tco_df, company_data)
        tagged = [a for a in all_alerts if a.get("company") == "TAGGED_CO"]
        assert len(tagged) >= 1
