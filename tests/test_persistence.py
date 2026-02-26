"""
test_persistence.py — Tests de services/persistence.py.

Utilise tmp_projects_dir (conftest) pour isoler les fichiers projet.
"""

from __future__ import annotations

from types import SimpleNamespace

import pytest

from services.persistence import delete_project, list_projects, load_project, save_project

# ---------------------------------------------------------------------------
# Helper
# ---------------------------------------------------------------------------


def make_state(**kwargs):
    """Crée un objet simulant st.session_state."""
    s = SimpleNamespace(**kwargs)
    s.get = lambda k, d=None: getattr(s, k, d)
    return s


# ---------------------------------------------------------------------------
# Tests validation nom projet
# ---------------------------------------------------------------------------


class TestProjectNameValidation:
    def test_rejects_path_traversal_dotdot(self, tmp_projects_dir):
        ok, msg = save_project("../evil", make_state())
        assert ok is False
        assert "invalide" in msg.lower()

    def test_rejects_forward_slash(self, tmp_projects_dir):
        ok, _ = save_project("path/evil", make_state())
        assert ok is False

    def test_rejects_backslash(self, tmp_projects_dir):
        ok, _ = save_project("path\\evil", make_state())
        assert ok is False

    def test_rejects_colon(self, tmp_projects_dir):
        ok, _ = save_project("C:evil", make_state())
        assert ok is False

    def test_rejects_empty_name(self, tmp_projects_dir):
        ok, _ = save_project("", make_state())
        assert ok is False

    def test_accepts_valid_name(self, tmp_projects_dir, minimal_tco_df):
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={},
            company_data={},
            step=1,
            all_alerts=[],
            tva_rate=0.20,
        )
        ok, _ = save_project("mon_projet_valide", state)
        assert ok is True


# ---------------------------------------------------------------------------
# Tests roundtrip save / load
# ---------------------------------------------------------------------------


class TestSaveLoadRoundtrip:
    def test_save_load_basic(self, tmp_projects_dir, minimal_tco_df):
        """Save puis load → DataFrame et step préservés."""
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={"project_info": {"projet": "TEST"}},
            company_data={},
            step=2,
            all_alerts=[],
            tva_rate=0.20,
        )
        ok, _ = save_project("test_basic", state)
        assert ok is True

        new_state = make_state()
        ok, msg = load_project("test_basic", new_state)
        assert ok is True, f"Chargement échoué : {msg}"
        assert new_state.tco_df is not None
        assert len(new_state.tco_df) == len(minimal_tco_df)
        assert new_state.step == 2

    def test_save_load_with_company_data(self, tmp_projects_dir, minimal_tco_df, minimal_dpgf_df):
        """Company data préservée après roundtrip."""
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={},
            company_data={
                "ACME": {
                    "dpgf_df": minimal_dpgf_df,
                    "parse_alerts": [],
                    "filename": "acme.xlsx",
                }
            },
            step=3,
            all_alerts=[],
            tva_rate=0.20,
        )
        ok, _ = save_project("test_company", state)
        assert ok is True

        new_state = make_state()
        ok, _ = load_project("test_company", new_state)
        assert ok is True
        assert "ACME" in new_state.company_data
        assert new_state.company_data["ACME"]["filename"] == "acme.xlsx"

    def test_decimal_serialized_as_numeric(self, tmp_projects_dir, minimal_tco_df):
        """Les Decimal sont sérialisés comme float (pas comme strings)."""
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={},
            company_data={},
            step=1,
            all_alerts=[],
            tva_rate=0.20,
        )
        save_project("test_decimal", state)

        new_state = make_state()
        load_project("test_decimal", new_state)

        # Les colonnes numériques ne doivent pas être des strings
        df = new_state.tco_df
        assert df is not None
        # La colonne Px_Tot_HT doit être numérique (float ou Decimal), pas string
        for val in df["Px_Tot_HT"]:
            assert not isinstance(val, str), f"Valeur string détectée : {val!r}"

    def test_load_nonexistent_project(self, tmp_projects_dir):
        """Chargement d'un projet inexistant → (False, msg)."""
        state = make_state()
        ok, msg = load_project("projet_inexistant_xyz", state)
        assert ok is False
        # Le message doit indiquer que le fichier n'existe pas
        assert "existe" in msg.lower() or "fichier" in msg.lower()

    def test_tva_rate_preserved(self, tmp_projects_dir, minimal_tco_df):
        """Le taux de TVA est préservé après save/load."""
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={},
            company_data={},
            step=1,
            all_alerts=[],
            tva_rate=0.055,
        )
        save_project("test_tva", state)
        new_state = make_state()
        load_project("test_tva", new_state)
        assert new_state.tva_rate == pytest.approx(0.055)


# ---------------------------------------------------------------------------
# Tests list_projects
# ---------------------------------------------------------------------------


class TestListProjects:
    def test_list_projects_sorted(self, tmp_projects_dir, minimal_tco_df):
        """list_projects() retourne les noms triés alphabétiquement."""
        for name in ["beta", "alpha", "gamma"]:
            state = make_state(
                tco_df=minimal_tco_df,
                merged_df=None,
                tco_meta={},
                company_data={},
                step=1,
                all_alerts=[],
                tva_rate=0.20,
            )
            save_project(name, state)

        names = list_projects()
        assert names == sorted(names)
        for name in ["alpha", "beta", "gamma"]:
            assert name in names

    def test_list_projects_empty_dir(self, tmp_projects_dir):
        """Aucun projet → liste vide."""
        assert list_projects() == []


# ---------------------------------------------------------------------------
# Tests delete_project
# ---------------------------------------------------------------------------


class TestDeleteProject:
    def test_delete_existing_project(self, tmp_projects_dir, minimal_tco_df):
        """delete_project sur un projet existant → True, absent de list_projects."""
        state = make_state(
            tco_df=minimal_tco_df,
            merged_df=None,
            tco_meta={},
            company_data={},
            step=1,
            all_alerts=[],
            tva_rate=0.20,
        )
        save_project("to_delete", state)
        assert "to_delete" in list_projects()

        result = delete_project("to_delete")
        assert result is True
        assert "to_delete" not in list_projects()

    def test_delete_nonexistent_project(self, tmp_projects_dir):
        """delete_project sur un projet inexistant → False."""
        result = delete_project("inexistant_xyz")
        assert result is False
