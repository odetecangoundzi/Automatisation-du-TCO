"""
test_file_validator.py — Tests des 5 fonctions de validation de fichiers.

Utilise io.BytesIO pour simuler des fichiers uploadés sans disque.
"""

from __future__ import annotations

import io

import openpyxl

from services.file_validator import (
    validate_extension,
    validate_magic_bytes,
    validate_size,
    validate_uploaded_file,
    validate_zip_bomb,
)

# ---------------------------------------------------------------------------
# Helpers et constantes
# ---------------------------------------------------------------------------

# Magic bytes standards
ZIP_MAGIC = b"PK\x03\x04"
XLS_MAGIC = b"\xd0\xcf\x11\xe0"
BAD_MAGIC = b"\x00\x01\x02\x03"


def make_uploaded_file(content: bytes, name: str = "test.xlsx") -> io.BytesIO:
    """Simule un UploadedFile Streamlit avec .name et .size."""
    buf = io.BytesIO(content)
    buf.name = name
    buf.size = len(content)
    return buf


def make_valid_xlsx_bytes() -> bytes:
    """Crée un vrai fichier .xlsx en mémoire via openpyxl."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.append(["Code", "Désignation", "Prix"])
    ws.append(["1.1", "Article test", 100])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# validate_extension
# ---------------------------------------------------------------------------


class TestValidateExtension:
    def test_xlsx_accepted(self):
        ok, msg = validate_extension("rapport.xlsx")
        assert ok is True
        assert msg == ""

    def test_xlsm_accepted(self):
        ok, _ = validate_extension("rapport.xlsm")
        assert ok is True

    def test_xls_accepted(self):
        ok, _ = validate_extension("rapport.xls")
        assert ok is True

    def test_xlsb_accepted(self):
        ok, _ = validate_extension("rapport.xlsb")
        assert ok is True

    def test_pdf_rejected(self):
        ok, msg = validate_extension("rapport.pdf")
        assert ok is False
        assert ".pdf" in msg

    def test_csv_rejected(self):
        ok, _ = validate_extension("data.csv")
        assert ok is False

    def test_case_insensitive_XLSX(self):
        """Extension en majuscules acceptée."""
        ok, _ = validate_extension("rapport.XLSX")
        assert ok is True

    def test_no_extension_rejected(self):
        ok, _ = validate_extension("rapport")
        assert ok is False


# ---------------------------------------------------------------------------
# validate_size
# ---------------------------------------------------------------------------


class TestValidateSize:
    def test_under_limit(self):
        ok, _ = validate_size(5 * 1024 * 1024, max_mb=20)
        assert ok is True

    def test_over_limit(self):
        ok, msg = validate_size(25 * 1024 * 1024, max_mb=20)
        assert ok is False
        assert "MB" in msg

    def test_exact_limit_accepted(self):
        """Exactement la limite → accepté."""
        ok, _ = validate_size(20 * 1024 * 1024, max_mb=20)
        assert ok is True

    def test_zero_size(self):
        ok, _ = validate_size(0, max_mb=20)
        assert ok is True


# ---------------------------------------------------------------------------
# validate_magic_bytes
# ---------------------------------------------------------------------------


class TestValidateMagicBytes:
    def test_zip_magic_valid(self):
        buf = make_uploaded_file(ZIP_MAGIC + b"\x00" * 100)
        ok, _ = validate_magic_bytes(buf)
        assert ok is True

    def test_xls_magic_valid(self):
        buf = make_uploaded_file(XLS_MAGIC + b"\x00" * 100, name="test.xls")
        ok, _ = validate_magic_bytes(buf)
        assert ok is True

    def test_invalid_magic_rejected(self):
        buf = make_uploaded_file(BAD_MAGIC + b"\x00" * 100)
        ok, msg = validate_magic_bytes(buf)
        assert ok is False
        assert "Excel" in msg or "fichier" in msg.lower()

    def test_too_small_rejected(self):
        buf = make_uploaded_file(b"\x00\x01")
        ok, msg = validate_magic_bytes(buf)
        assert ok is False
        assert "petit" in msg.lower() or "vide" in msg.lower()

    def test_seek_reset_after_validation(self):
        """Après validate_magic_bytes, la position de lecture est remise à 0."""
        buf = make_uploaded_file(ZIP_MAGIC + b"\xab" * 100)
        validate_magic_bytes(buf)
        # La position doit être à 0 pour pouvoir lire à nouveau
        header = buf.read(4)
        assert header == ZIP_MAGIC


# ---------------------------------------------------------------------------
# validate_zip_bomb
# ---------------------------------------------------------------------------


class TestValidateZipBomb:
    def test_normal_xlsx_passes(self):
        """Un vrai .xlsx avec ratio de compression normal → (True, '')."""
        content = make_valid_xlsx_bytes()
        buf = make_uploaded_file(content)
        ok, _ = validate_zip_bomb(buf)
        assert ok is True

    def test_non_zip_file_ignored(self):
        """Un fichier non-ZIP (XLS OLE) → (True, '') sans exception."""
        buf = make_uploaded_file(XLS_MAGIC + b"\x00" * 200, name="test.xls")
        ok, _ = validate_zip_bomb(buf)
        assert ok is True

    def test_seek_reset_after_validation(self):
        """Après validate_zip_bomb, la position est remise à 0."""
        content = make_valid_xlsx_bytes()
        buf = make_uploaded_file(content)
        validate_zip_bomb(buf)
        buf.seek(0)
        header = buf.read(4)
        assert header == ZIP_MAGIC


# ---------------------------------------------------------------------------
# validate_uploaded_file (composite)
# ---------------------------------------------------------------------------


class TestValidateUploadedFile:
    def test_valid_xlsx_file(self):
        """Vrai .xlsx valide → (True, '')."""
        content = make_valid_xlsx_bytes()
        buf = make_uploaded_file(content, name="rapport.xlsx")
        ok, msg = validate_uploaded_file(buf)
        assert ok is True, f"Attendu True mais obtenu False : {msg}"

    def test_bad_extension_short_circuits(self):
        """Extension .pdf avec magic bytes ZIP → rejeté dès l'extension."""
        buf = make_uploaded_file(ZIP_MAGIC + b"\x00" * 100, name="rapport.pdf")
        ok, _ = validate_uploaded_file(buf)
        assert ok is False

    def test_bad_magic_after_valid_extension(self):
        """Extension .xlsx correcte mais magic bytes invalides → rejeté."""
        buf = make_uploaded_file(BAD_MAGIC + b"\x00" * 100, name="rapport.xlsx")
        ok, msg = validate_uploaded_file(buf)
        assert ok is False
        assert "Excel" in msg or "fichier" in msg.lower()

    def test_oversized_file_rejected(self):
        """Fichier > 20 MB → rejeté à l'étape taille."""
        content = make_valid_xlsx_bytes()
        buf = io.BytesIO(content)
        buf.name = "rapport.xlsx"
        buf.size = 25 * 1024 * 1024  # simuler taille artificielle > 20 MB
        ok, msg = validate_uploaded_file(buf, max_mb=20)
        assert ok is False
        assert "MB" in msg
