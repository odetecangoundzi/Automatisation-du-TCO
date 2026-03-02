"""
conftest.py — Fixtures partagées entre tous les fichiers de test.

Ne jamais importer depuis app.py (contient Streamlit).
"""

from __future__ import annotations

import io
from decimal import Decimal

import openpyxl
import pandas as pd
import pytest

# ---------------------------------------------------------------------------
# Helpers (fonctions utilitaires, pas des fixtures pytest)
# ---------------------------------------------------------------------------


def make_excel_bytes(rows: list[list], header: list | None = None) -> io.BytesIO:
    """Crée un fichier .xlsx en mémoire avec les lignes données."""
    wb = openpyxl.Workbook()
    ws = wb.active
    if header:
        ws.append(header)
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


class SimpleState:
    """Simule st.session_state pour les tests de persistence."""

    def get(self, key: str, default=None):
        return getattr(self, key, default)


# ---------------------------------------------------------------------------
# DataFrames inline
# ---------------------------------------------------------------------------


@pytest.fixture
def minimal_tco_df() -> pd.DataFrame:
    """TCO minimal : 1 section + 1 article + 1 recap."""
    return pd.DataFrame(
        [
            {
                "Code": "1",
                "Désignation": "Section A",
                "Qu.": Decimal("0"),
                "U": "",
                "Px_U_HT": Decimal("0"),
                "Px_Tot_HT": Decimal("0"),
                "Entete": "Bd_01_Bord",
                "row_type": "section_header",
                "original_row": 2,
                "parent_code": "",
            },
            {
                "Code": "1.1",
                "Désignation": "Article un",
                "Qu.": Decimal("10"),
                "U": "u",
                "Px_U_HT": Decimal("20"),
                "Px_Tot_HT": Decimal("200"),
                "Entete": "Ouv_01_Art",
                "row_type": "article",
                "original_row": 3,
                "parent_code": "",
            },
            {
                "Code": "",
                "Désignation": "Total section A",
                "Qu.": Decimal("0"),
                "U": "",
                "Px_U_HT": Decimal("0"),
                "Px_Tot_HT": Decimal("0"),
                "Entete": "Bord_01_Recap",
                "row_type": "recap",
                "original_row": 4,
                "parent_code": "1",
            },
        ]
    )


@pytest.fixture
def minimal_dpgf_df() -> pd.DataFrame:
    """DPGF minimal : 1 article matchant le code 1.1 du TCO."""
    return pd.DataFrame(
        [
            {
                "Code": "1.1",
                "Désignation": "Article un DPGF",
                "Qu.": Decimal("10"),
                "U": "u",
                "Px_U_HT": Decimal("25"),
                "Px_Tot_HT": Decimal("250"),
                "Commentaire": "",
                "Entete": "Ouv_01_Art",
                "row_type": "article",
                "original_row": 2,
                "parent_code": "",
            }
        ]
    )


# ---------------------------------------------------------------------------
# Isolation PROJECTS_DIR pour tests persistence
# ---------------------------------------------------------------------------


@pytest.fixture
def tmp_projects_dir(tmp_path, monkeypatch):
    """Redirige PROJECTS_DIR vers un répertoire temporaire isolé."""
    projects = tmp_path / "projects"
    projects.mkdir(parents=True, exist_ok=True)
    monkeypatch.setattr("services.persistence.PROJECTS_DIR", str(projects))
    return projects
