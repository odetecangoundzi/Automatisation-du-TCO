"""
test_parser_dpgf_alerts.py — Tests des 5 types d'alertes du parser DPGF.

Chaque test crée un fichier Excel minimal via openpyxl + tmp_path qui
déclenche exactement une alerte cible.
"""

from decimal import Decimal

import openpyxl
import pytest

from core.parser_dpgf import parse_dpgf

# ---------------------------------------------------------------------------
# Fixture factory
# ---------------------------------------------------------------------------

HEADER = ["Code", "Désignation", "Qu.", "U", "Px U", "Px Tot"]


@pytest.fixture
def dpgf_file_factory(tmp_path):
    """Retourne une fonction créant un .xlsx minimal avec le header standard."""

    def _make(rows: list[list], filename: str = "dpgf_test.xlsx") -> str:
        fpath = tmp_path / filename
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(HEADER)
        for row in rows:
            ws.append(row)
        wb.save(fpath)
        return str(fpath)

    return _make


# ---------------------------------------------------------------------------
# Alerte 1 : Codes dupliqués
# ---------------------------------------------------------------------------


class TestDuplicateCodes:
    def test_duplicate_codes_suffix_dup02(self, dpgf_file_factory):
        """2 articles avec le même code → _DUP02 sur le second."""
        path = dpgf_file_factory(
            [
                ["1.1", "Article A", 10, "u", 20, 200],
                ["1.1", "Article A bis", 5, "u", 30, 150],
            ]
        )
        df, alerts = parse_dpgf(path)

        codes = df["Code"].tolist()
        assert "1.1" in codes
        assert "1.1_DUP02" in codes

        assert "Code_source" in df.columns
        # La source originale est préservée pour les deux occurrences
        assert df[df["Code"] == "1.1"]["Code_source"].iloc[0] == "1.1"
        assert df[df["Code"] == "1.1_DUP02"]["Code_source"].iloc[0] == "1.1"

        dup_alerts = [a for a in alerts if "dupliqué" in a["message"].lower()]
        assert len(dup_alerts) >= 1
        assert dup_alerts[0]["type"] == "warning"

    def test_three_duplicate_codes(self, dpgf_file_factory):
        """3 articles avec le même code → _DUP02, _DUP03."""
        path = dpgf_file_factory(
            [
                ["2.1", "Article X", 1, "u", 100, 100],
                ["2.1", "Article X bis", 2, "u", 100, 200],
                ["2.1", "Article X ter", 3, "u", 100, 300],
            ]
        )
        df, alerts = parse_dpgf(path)

        codes = df["Code"].tolist()
        assert "2.1" in codes
        assert "2.1_DUP02" in codes
        assert "2.1_DUP03" in codes

    def test_no_code_source_without_duplicates(self, dpgf_file_factory):
        """Sans doublon, la colonne Code_source n'est PAS créée."""
        path = dpgf_file_factory(
            [
                ["1.1", "Article unique", 10, "u", 20, 200],
            ]
        )
        df, alerts = parse_dpgf(path)
        assert "Code_source" not in df.columns


# ---------------------------------------------------------------------------
# Alerte 2 : Texte dans champs numériques
# ---------------------------------------------------------------------------


class TestTextInNumericFields:
    # Note : les alertes texte/mot-clé ne sont générées que pour row_type == "article".
    # classify_row classifie "article" si len(code.split(".")) >= 4 OU si has_price=True.
    # Avec un mot-clé dans Px U, has_price=False → il faut un code à 4 segments.

    def test_keyword_sans_objet_info_blue(self, dpgf_file_factory):
        """'SANS OBJET' dans Px U → alerte info bleue, Commentaire='so'."""
        # Code 4 segments → article (indépendamment de has_price)
        path = dpgf_file_factory(
            [
                ["1.1.1.1", "Article SO", 10, "u", "SANS OBJET", 0],
            ]
        )
        df, alerts = parse_dpgf(path)

        info_alerts = [a for a in alerts if a["type"] == "info" and a["color"] == "blue"]
        assert len(info_alerts) >= 1
        assert "Mot-clé détecté" in info_alerts[0]["message"]

        row = df[df["Code"] == "1.1.1.1"].iloc[0]
        assert row["Px_U_HT"] == Decimal("0.0")
        assert "so" in row["Commentaire"]

    def test_keyword_compris_in_quantity(self, dpgf_file_factory):
        """'COMPRIS' dans Qu. → commentaire contient 'compris'."""
        path = dpgf_file_factory(
            [
                ["1.1.1.1", "Article compris", "COMPRIS", "u", 50, 0],
            ]
        )
        df, alerts = parse_dpgf(path)

        info_alerts = [a for a in alerts if a["type"] == "info"]
        assert len(info_alerts) >= 1
        row = df[df["Code"] == "1.1.1.1"].iloc[0]
        assert "compris" in row["Commentaire"]

    def test_text_non_keyword_warning_yellow(self, dpgf_file_factory):
        """Texte libre non-keyword → alerte warning jaune."""
        path = dpgf_file_factory(
            [
                ["1.1.1.1", "Article bordereau", 10, "u", "Voir bordereau page 12", 0],
            ]
        )
        df, alerts = parse_dpgf(path)

        warn_alerts = [a for a in alerts if a["type"] == "warning" and a["color"] == "yellow"]
        assert len(warn_alerts) >= 1
        assert "Texte dans champ numérique" in warn_alerts[0]["message"]

    def test_keyword_inclus_info(self, dpgf_file_factory):
        """'INCLUS' dans Qu. → alerte info bleue.

        Note : "nc" est substring de "inclus" et apparaît avant dans KEYWORDS,
        donc l'abréviation réelle dans Commentaire est "nc" (non-chiffré).
        """
        path = dpgf_file_factory(
            [
                ["1.1.1.1", "Article inclus", "INCLUS", "u", 0, 0],
            ]
        )
        df, alerts = parse_dpgf(path)

        info_alerts = [a for a in alerts if a["type"] == "info" and a["color"] == "blue"]
        assert len(info_alerts) >= 1
        row = df[df["Code"] == "1.1.1.1"].iloc[0]
        # "nc" est matchée avant "inclus" car c'est un substring de "INCLUS"
        assert row["Commentaire"] != ""


# ---------------------------------------------------------------------------
# Alerte 3 : Total incohérent
# ---------------------------------------------------------------------------


class TestTotalCoherence:
    def test_total_incoherence_error_red(self, dpgf_file_factory):
        """Qu × PU ≠ Total (écart > 0.10€ et > 0.1%) → alerte error rouge."""
        # 10 × 20 = 200 mais on met 300 (écart 100€)
        path = dpgf_file_factory(
            [
                ["1.1", "Article incohérent", 10, "u", 20, 300],
            ]
        )
        df, alerts = parse_dpgf(path)

        error_alerts = [a for a in alerts if a["type"] == "error" and a["color"] == "red"]
        assert len(error_alerts) >= 1
        assert "Total incohérent" in error_alerts[0]["message"]
        assert "1.1" == error_alerts[0]["code"]

    def test_no_alert_within_abs_tolerance(self, dpgf_file_factory):
        """Écart < 0.10€ → aucune alerte de total incohérent."""
        # 10 × 20.009 = 200.09 ≈ 200 (écart 0.09 < 0.10)
        path = dpgf_file_factory(
            [
                ["1.1", "Article tolérance abs", 10, "u", 20.009, 200],
            ]
        )
        df, alerts = parse_dpgf(path)

        error_alerts = [a for a in alerts if "Total incohérent" in a.get("message", "")]
        assert len(error_alerts) == 0

    def test_no_alert_within_rel_tolerance(self, dpgf_file_factory):
        """Écart relatif < 0.1% → aucune alerte de total incohérent."""
        # 1000 × 100.0001 = 100000.1 ≈ 100000 (rel = 0.000001 < 0.001)
        path = dpgf_file_factory(
            [
                ["1.1", "Article tolérance rel", 1000, "u", 100.0001, 100000],
            ]
        )
        df, alerts = parse_dpgf(path)

        error_alerts = [a for a in alerts if "Total incohérent" in a.get("message", "")]
        assert len(error_alerts) == 0


# ---------------------------------------------------------------------------
# Alerte 4 : Unité manquante
# ---------------------------------------------------------------------------


class TestMissingUnit:
    def test_missing_unit_warning_orange(self, dpgf_file_factory):
        """Unité vide avec PxTot > 0 → alerte warning orange."""
        path = dpgf_file_factory(
            [
                ["1.1", "Article sans unité", 1, "", 500, 500],
            ]
        )
        df, alerts = parse_dpgf(path)

        unit_alerts = [a for a in alerts if "Unité manquante" in a.get("message", "")]
        assert len(unit_alerts) >= 1
        assert unit_alerts[0]["type"] == "warning"
        assert unit_alerts[0]["color"] == "orange"

    def test_no_alert_missing_unit_when_total_zero(self, dpgf_file_factory):
        """Unité vide mais PxTot = 0 → pas d'alerte unité manquante."""
        path = dpgf_file_factory(
            [
                ["1.1", "Article gratuit", 0, "", 0, 0],
            ]
        )
        df, alerts = parse_dpgf(path)

        unit_alerts = [a for a in alerts if "Unité manquante" in a.get("message", "")]
        assert len(unit_alerts) == 0


# ---------------------------------------------------------------------------
# Normalisation code float Excel
# ---------------------------------------------------------------------------


class TestFloatCodeNormalization:
    def test_float_code_normalized_to_string(self, dpgf_file_factory):
        """Cellule Code numérique 1.1 (float Excel) → df["Code"] == "1.1"."""
        path = dpgf_file_factory(
            [
                # openpyxl écrit 1.1 comme valeur numérique (comme Excel ferait avec "01.1")
                [1.1, "Article float code", 10, "u", 20, 200],
            ]
        )
        df, alerts = parse_dpgf(path)

        assert not df.empty
        # Le code doit être normalisé en string "1.1"
        assert df.iloc[0]["Code"] == "1.1"
