"""
test_utils.py — Tests unitaires pour core/utils.py

Couvre : classify_row, find_header_row, find_column_index.
"""

import pandas as pd
import pytest

from core.utils import COL_NOT_FOUND, classify_row, find_column_index, find_header_row

# ---------------------------------------------------------------------------
# classify_row
# ---------------------------------------------------------------------------


class TestClassifyRow:
    """Tests pour classify_row — 6 chemins de décision."""

    # Chemin 1 : Entete (priorité haute)

    def test_entete_recapbord_recap_summary(self):
        """RecapBord_xxx → recap_summary (priorité absolue)."""
        assert classify_row("1", "Récap", "RecapBord_01") == "recap_summary"

    def test_entete_lignes_tot_total_line(self):
        """LignesTot_xxx → total_line."""
        assert classify_row("", "Montant", "LignesTot_01") == "total_line"

    def test_entete_bord_recap_recap(self):
        """Bord + Recap dans entete → recap."""
        assert classify_row("", "Total lot", "Bord_01_Recap") == "recap"

    def test_entete_bd_bord_section_header(self):
        """Bd_ + Bord → section_header."""
        assert classify_row("1", "Lot 1", "Bd_01_Bord") == "section_header"

    def test_entete_niv1_sub_section(self):
        """_Niv1 → sub_section."""
        assert classify_row("1.1.1", "Sous-section", "Ouv_01_Niv1") == "sub_section"

    def test_entete_niv2_sub_section(self):
        """_Niv2 → sub_section."""
        assert classify_row("1.1.2", "Sous-section", "Ouv_01_Niv2") == "sub_section"

    def test_entete_art_article(self):
        """_Art → article."""
        assert classify_row("1.1", "Article", "Ouv_01_Art") == "article"

    # Chemin 2 : Désignation (fallback totaux)

    def test_desig_montant_ht_total_line(self):
        """'montant ht' dans désignation → total_line."""
        assert classify_row("", "Montant HT", "") == "total_line"

    def test_desig_montant_ttc_total_line(self):
        """'montant ttc' dans désignation → total_line."""
        assert classify_row("", "Montant TTC", "") == "total_line"

    def test_desig_total_section_recap(self):
        """'total' + 'section' → recap."""
        assert classify_row("", "Total section 01", "") == "recap"

    # Chemin 3 : Priorité prix

    def test_has_price_with_code_article(self):
        """has_price=True + code renseigné → article."""
        assert classify_row("1.1", "Travaux", "", has_price=True) == "article"

    def test_has_price_without_code_article(self):
        """has_price=True sans code (ex: ERTIE&FILS) → article (fallback final)."""
        assert classify_row("", "Travaux sans code", "", has_price=True) == "article"

    # Chemin 4 : Structure du code

    def test_code_one_part_section_header(self):
        """Code "01" (1 partie) → section_header."""
        assert classify_row("01", "", "") == "section_header"

    def test_code_two_parts_section_header(self):
        """Code "01.1" (2 parties) → section_header."""
        assert classify_row("01.1", "", "") == "section_header"

    def test_code_three_parts_sub_section(self):
        """Code "01.1.2" (3 parties) → sub_section."""
        assert classify_row("01.1.2", "", "") == "sub_section"

    def test_code_four_parts_article(self):
        """Code "01.1.1.1" (4 parties) → article."""
        assert classify_row("01.1.1.1", "", "") == "article"

    def test_code_starts_with_total_total_text(self):
        """Code commençant par 'total' → total_text."""
        assert classify_row("Total lot 01", "", "") == "total_text"

    # Chemin 5 : Vide

    def test_empty_code_and_desig_empty(self):
        """Ni code ni désignation → empty."""
        assert classify_row("", "", "") == "empty"

    # Chemin 6 : other

    def test_other_fallback(self):
        """Désignation seule sans prix → other."""
        assert classify_row("", "Description libre", "") == "other"


# ---------------------------------------------------------------------------
# find_header_row
# ---------------------------------------------------------------------------


class TestFindHeaderRow:
    """Tests pour find_header_row."""

    def test_standard_header_at_row_0(self):
        """Header standard en ligne 0."""
        df = pd.DataFrame([["Code", "Désignation", "Qu.", "U", "Px U", "Px Tot"]])
        assert find_header_row(df) == 0

    def test_header_after_prefix_rows(self):
        """Header en ligne 3 (après 3 lignes de métadonnées)."""
        data = [
            ["Projet XYZ", None, None],
            ["Région Sud", None, None],
            [None, None, None],
            ["Code", "Désignation", "Px U"],
        ]
        df = pd.DataFrame(data)
        assert find_header_row(df) == 3

    def test_synonym_numero_libelle(self):
        """Synonymes N° + Libellé acceptés."""
        df = pd.DataFrame([["N°", "Libellé", "Qu.", "U", "P.U", "Total HT"]])
        assert find_header_row(df) == 0

    def test_raises_value_error_when_not_found(self):
        """ValueError si aucun header reconnaissable."""
        df = pd.DataFrame(
            [
                ["Coucou", "Monde"],
                ["123", "456"],
            ]
        )
        with pytest.raises(ValueError, match="Impossible de trouver"):
            find_header_row(df)

    def test_fallback_desig_plus_price_without_code(self):
        """Fallback : Désignation + colonne prix sans colonne Code (cas ERTIE&FILS)."""
        df = pd.DataFrame([["Désignation", "Qu.", "U", "P.U", "Total HT"]])
        assert find_header_row(df) == 0


# ---------------------------------------------------------------------------
# find_column_index
# ---------------------------------------------------------------------------


class TestFindColumnIndex:
    """Tests pour find_column_index."""

    def _make_df(self, columns: list[str]) -> pd.DataFrame:
        return pd.DataFrame(columns=columns)

    def test_exact_match(self):
        """Match exact sur 'code'."""
        df = self._make_df(["code", "désignation", "qu.", "u", "px u", "px tot"])
        assert find_column_index(df, ["code"]) == 0

    def test_substring_match(self):
        """Sous-chaîne 'prix u' dans 'Prix Unitaire HT'."""
        df = self._make_df(["Code", "Désignation", "Prix Unitaire HT"])
        assert find_column_index(df, ["prix u"]) == 2

    def test_not_found_returns_col_not_found(self):
        """Absent + default_idx=None → COL_NOT_FOUND."""
        df = self._make_df(["Code", "Désignation"])
        assert find_column_index(df, ["entete"]) == COL_NOT_FOUND

    def test_not_found_returns_default(self):
        """Absent + default_idx fourni → default_idx."""
        df = self._make_df(["Code", "Désignation"])
        assert find_column_index(df, ["entete"], default_idx=5) == 5

    def test_single_char_keyword_exact_only(self):
        """Keyword 1 char 'u' ne doit PAS matcher 'qu. ent.'."""
        df = self._make_df(["qu. ent.", "unité"])
        # "u" doit matcher "unité" (substring) mais pas "qu. ent." (exact 1-char)
        result = find_column_index(df, ["u"])
        # "unité" contient "u" mais keyword 1-char = exact, col_base "unité".rstrip(". ") = "unité"
        # Donc "u" != "unité" → pas de match sauf via substring? Non :
        # len("u") == 1 → exact only : kw_base="u", col_base doit être "u"
        # "qu. ent." → col_base = "qu. ent" ≠ "u" ✓
        # "unité" → col_base = "unité" ≠ "u" → COL_NOT_FOUND
        assert result == COL_NOT_FOUND

    def test_multi_keyword_first_match(self):
        """Plusieurs keywords : retourne le premier match."""
        df = self._make_df(["code", "désignation", "quantité"])
        # "qu." → col_base "qu." rstrip → "qu", kw_base "qu" → match "quantité"? Non.
        # "quantité" contient "qu." ? kw="qu." len=3 > 1 → substring : "qu." in "quantité" → non
        # Cherchons "quantité" directement
        assert find_column_index(df, ["qte", "quantité", "qt"]) == 2
