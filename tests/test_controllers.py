"""
test_controllers.py — Tests unitaires de app/controllers.py.

Couvre : validate_company_name, normalize_filename, rebuild_merged_tco.
Aucune dépendance Streamlit — tout est pur Python.
"""

from __future__ import annotations

from app.controllers import normalize_filename, rebuild_merged_tco, validate_company_name
from config import COMPANY_NAME_MAX_LEN

# ---------------------------------------------------------------------------
# Tests validate_company_name
# ---------------------------------------------------------------------------


class TestValidateCompanyName:
    def test_valid_name(self):
        """Nom valide → (name_clean, None)."""
        name, err = validate_company_name("ACME SA")
        assert name == "ACME SA"
        assert err is None

    def test_valid_name_with_allowed_chars(self):
        """Nom avec tirets, points et parenthèses → valide."""
        name, err = validate_company_name("BTP-Nord (06)")
        assert name == "BTP-Nord (06)"
        assert err is None

    def test_strips_surrounding_whitespace(self):
        """Les espaces en début/fin sont supprimés."""
        name, err = validate_company_name("  ACME  ")
        assert name == "ACME"
        assert err is None

    def test_empty_name(self):
        """Nom vide → (None, message d'erreur)."""
        name, err = validate_company_name("")
        assert name is None
        assert err is not None
        assert len(err) > 0

    def test_whitespace_only(self):
        """Nom avec que des espaces → équivalent vide."""
        name, err = validate_company_name("   ")
        assert name is None
        assert err is not None

    def test_too_long_name(self):
        """Nom dépassant COMPANY_NAME_MAX_LEN → (None, message d'erreur)."""
        long_name = "A" * (COMPANY_NAME_MAX_LEN + 1)
        name, err = validate_company_name(long_name)
        assert name is None
        assert err is not None
        assert str(COMPANY_NAME_MAX_LEN) in err

    def test_exact_max_length_accepted(self):
        """Nom de longueur exacte COMPANY_NAME_MAX_LEN → valide."""
        exact_name = "A" * COMPANY_NAME_MAX_LEN
        name, err = validate_company_name(exact_name)
        assert name == exact_name
        assert err is None

    def test_special_chars_ampersand(self):
        """& interdit (risque HTML/SQL) → (None, message d'erreur)."""
        name, err = validate_company_name("ACME&SA")
        assert name is None
        assert err is not None

    def test_special_chars_quote(self):
        """' interdit → (None, message d'erreur)."""
        name, err = validate_company_name("L'Entreprise")
        assert name is None
        assert err is not None

    def test_special_chars_semicolon(self):
        """; interdit → (None, message d'erreur)."""
        name, err = validate_company_name("ACME;SA")
        assert name is None
        assert err is not None


# ---------------------------------------------------------------------------
# Tests normalize_filename
# ---------------------------------------------------------------------------


class TestNormalizeFilename:
    def test_basic_normalization(self):
        """Espaces et accents → underscores, majuscules."""
        result = normalize_filename("Chantier Bordeaux 2026")
        assert result == "CHANTIER_BORDEAUX_2026"

    def test_to_uppercase(self):
        """Toujours converti en majuscules."""
        result = normalize_filename("hello")
        assert result == "HELLO"

    def test_special_chars_replaced(self):
        """Caractères non alphanumérique → underscore."""
        result = normalize_filename("ACME-BTP (Nord)")
        assert "ACME" in result
        assert "BTP" in result
        assert "NORD" in result
        assert "-" not in result
        assert "(" not in result

    def test_collapses_consecutive_underscores(self):
        """Underscores consécutifs → un seul underscore."""
        result = normalize_filename("A  B")  # 2 espaces → _
        assert "__" not in result
        assert result == "A_B"

    def test_strips_leading_trailing_underscores(self):
        """Pas d'underscore en début/fin."""
        result = normalize_filename("  ACME  ")
        assert not result.startswith("_")
        assert not result.endswith("_")
        assert result == "ACME"

    def test_numbers_preserved(self):
        """Les chiffres sont conservés."""
        result = normalize_filename("LOT 01")
        assert result == "LOT_01"

    def test_empty_string(self):
        """Chaîne vide → chaîne vide."""
        result = normalize_filename("")
        assert result == ""


# ---------------------------------------------------------------------------
# Tests rebuild_merged_tco
# ---------------------------------------------------------------------------


class TestRebuildMergedTCO:
    def test_full_rebuild_adds_company_columns(self, minimal_tco_df, minimal_dpgf_df):
        """Reconstruction complète → colonnes {ACME}_Px_Tot_HT présentes."""
        company_data = {
            "ACME": {
                "dpgf_df": minimal_dpgf_df,
                "parse_alerts": [],
                "filename": "acme.xlsx",
            }
        }
        merged_df, alerts = rebuild_merged_tco(minimal_tco_df, company_data, 0.20)

        assert merged_df is not None
        assert "ACME_Px_Tot_HT" in merged_df.columns
        assert "ACME_Px_U_HT" in merged_df.columns

    def test_full_rebuild_returns_list_of_alerts(self, minimal_tco_df, minimal_dpgf_df):
        """rebuild_merged_tco retourne toujours une liste d'alertes (peut être vide)."""
        company_data = {
            "ACME": {
                "dpgf_df": minimal_dpgf_df,
                "parse_alerts": [],
                "filename": "acme.xlsx",
            }
        }
        merged_df, alerts = rebuild_merged_tco(minimal_tco_df, company_data, 0.20)

        assert isinstance(alerts, list)

    def test_no_companies_returns_base_df(self, minimal_tco_df):
        """Aucune entreprise → DataFrame de base sans colonnes _Px_Tot_HT."""
        merged_df, alerts = rebuild_merged_tco(minimal_tco_df, {}, 0.20)

        assert merged_df is not None
        assert isinstance(alerts, list)
        # Aucune colonne entreprise
        extra_cols = [c for c in merged_df.columns if "_Px_Tot_HT" in c]
        assert extra_cols == []

    def test_incremental_merge_adds_columns(self, minimal_tco_df, minimal_dpgf_df):
        """Fusion incrémentale (new_companies fourni) → colonnes ACME_* ajoutées."""
        # D'abord une reconstruction complète sans entreprises
        base_merged, _ = rebuild_merged_tco(minimal_tco_df, {}, 0.20)

        # Puis fusion incrémentale avec ACME
        company_data = {
            "ACME": {
                "dpgf_df": minimal_dpgf_df,
                "parse_alerts": [],
                "filename": "acme.xlsx",
            }
        }
        merged_df, new_alerts = rebuild_merged_tco(
            minimal_tco_df,
            company_data,
            0.20,
            merged_df=base_merged,
            new_companies=["ACME"],
        )

        assert "ACME_Px_Tot_HT" in merged_df.columns
        assert isinstance(new_alerts, list)

    def test_incremental_only_returns_new_alerts(self, minimal_tco_df, minimal_dpgf_df):
        """Le chemin incrémental ne retourne que les alertes de la nouvelle entreprise."""
        company_data_a = {
            "ACME": {
                "dpgf_df": minimal_dpgf_df,
                "parse_alerts": [{"type": "info", "message": "alerte ACME"}],
                "filename": "acme.xlsx",
            }
        }
        # Full rebuild pour avoir un base merged_df
        base_merged, _ = rebuild_merged_tco(minimal_tco_df, company_data_a, 0.20)

        # Second company BETA
        dpgf_beta = minimal_dpgf_df.copy()
        company_data_b = {
            "BETA": {
                "dpgf_df": dpgf_beta,
                "parse_alerts": [{"type": "warning", "message": "alerte BETA"}],
                "filename": "beta.xlsx",
            }
        }
        all_company_data = {**company_data_a, **company_data_b}
        _, new_alerts = rebuild_merged_tco(
            minimal_tco_df,
            all_company_data,
            0.20,
            merged_df=base_merged,
            new_companies=["BETA"],
        )

        # Seules les alertes BETA sont retournées
        companies_in_alerts = {a.get("company") for a in new_alerts}
        assert "BETA" in companies_in_alerts
        assert "ACME" not in companies_in_alerts

    def test_incremental_without_merged_df_falls_back_to_full(
        self, minimal_tco_df, minimal_dpgf_df
    ):
        """new_companies sans merged_df → reconstruction complète (fallback)."""
        company_data = {
            "ACME": {
                "dpgf_df": minimal_dpgf_df,
                "parse_alerts": [],
                "filename": "acme.xlsx",
            }
        }
        # merged_df=None → reconstruction complète malgré new_companies
        merged_df, alerts = rebuild_merged_tco(
            minimal_tco_df,
            company_data,
            0.20,
            merged_df=None,
            new_companies=["ACME"],
        )

        assert "ACME_Px_Tot_HT" in merged_df.columns
