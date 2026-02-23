"""
file_validator.py — Validation sécurisée des fichiers uploadés.

Vérifie : extension, taille, magic bytes (ZIP/XLSX), structure interne.
"""

import os
import openpyxl
from logger import get_logger

log = get_logger(__name__)

# Magic bytes pour ZIP (XLSX/XLSM/XLSB) et OLE (XLS)
ZIP_MAGIC = b"PK\x03\x04"
XLS_MAGIC = b"\xd0\xcf\x11\xe0"

ALLOWED_EXTENSIONS: set[str] = {".xlsx", ".xlsm", ".xls", ".xlsb"}


def validate_extension(filename: str, allowed: set[str] | None = None) -> tuple[bool, str]:
    """Vérifie que l'extension du fichier est autorisée."""
    allowed = allowed or ALLOWED_EXTENSIONS
    ext = os.path.splitext(filename)[1].lower()
    if ext not in allowed:
        return False, f"Format non accepté : {ext}. Seuls {', '.join(allowed)} sont autorisés."
    return True, ""


def validate_size(file_size_bytes: int, max_mb: int = 20) -> tuple[bool, str]:
    """Vérifie que le fichier ne dépasse pas la taille max."""
    size_mb = file_size_bytes / (1024 * 1024)
    if size_mb > max_mb:
        return False, f"Fichier trop volumineux ({size_mb:.1f} MB > {max_mb} MB)."
    return True, ""


def validate_magic_bytes(uploaded_file) -> tuple[bool, str]:
    """
    Vérifie que le fichier commence par les magic bytes ZIP/XLSX/XLSM/XLSB ou OLE/XLS.

    Args:
        uploaded_file: objet fichier avec .read() et .seek()

    Returns:
        (is_valid, error_message)
    """
    uploaded_file.seek(0)
    header = uploaded_file.read(4)
    uploaded_file.seek(0)

    if len(header) < 4:
        return False, "Fichier trop petit ou vide."
    
    if header == ZIP_MAGIC:
        return True, ""
    if header == XLS_MAGIC:
        return True, ""
        
    return False, "Le contenu du fichier ne correspond pas à un fichier Excel valide."


def validate_excel_structure(uploaded_file) -> tuple[bool, str]:
    """
    Tente d'ouvrir le fichier avec pandas pour vérifier sa structure.

    Args:
        uploaded_file: objet fichier avec .read() et .seek()

    Returns:
        (is_valid, error_message)
    """
    import pandas as pd
    try:
        uploaded_file.seek(0)
        ext = os.path.splitext(uploaded_file.name)[1].lower()
        
        # Détermination de l'engine
        engine = None
        if ext == ".xls": engine = "xlrd"
        elif ext == ".xlsb": engine = "pyxlsb"
        elif ext in [".xlsx", ".xlsm"]: engine = "openpyxl"
        
        # Test de lecture rapide (juste les headers ou une petite partie)
        pd.read_excel(uploaded_file, engine=engine, nrows=1)
        
        uploaded_file.seek(0)
        return True, ""
    except Exception as e:
        log.warning("Fichier Excel illisible (%s) : %s", engine, type(e).__name__)
        uploaded_file.seek(0)
        return False, f"Le fichier Excel est corrompu ou illisible ({type(e).__name__})."


def validate_uploaded_file(
    uploaded_file,
    max_mb: int = 20,
    allowed_extensions: set[str] | None = None,
    check_structure: bool = False,
) -> tuple[bool, str]:
    """
    Validation complète d'un fichier uploadé.

    Vérifie dans l'ordre : extension, taille, magic bytes, et optionnellement
    la structure interne du fichier Excel.

    Args:
        uploaded_file: objet UploadedFile Streamlit
        max_mb: taille maximale en mégaoctets
        allowed_extensions: extensions autorisées
        check_structure: si True, vérifie aussi la structure Excel

    Returns:
        (is_valid, error_message). Si valide, error_message est vide.
    """
    # 1. Extension
    ok, err = validate_extension(uploaded_file.name, allowed_extensions)
    if not ok:
        log.warning("Upload refusé (extension) : %s", uploaded_file.name)
        return False, err

    # 2. Taille
    ok, err = validate_size(uploaded_file.size, max_mb)
    if not ok:
        log.warning("Upload refusé (taille) : %d octets", uploaded_file.size)
        return False, err

    # 3. Magic bytes
    ok, err = validate_magic_bytes(uploaded_file)
    if not ok:
        log.warning("Upload refusé (magic bytes invalides) : %s", uploaded_file.name)
        return False, err

    # 4. Structure (optionnel, plus coûteux)
    if check_structure:
        ok, err = validate_excel_structure(uploaded_file)
        if not ok:
            return False, err

    return True, ""

