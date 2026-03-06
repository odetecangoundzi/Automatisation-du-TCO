# -*- mode: python ; coding: utf-8 -*-
"""
TCO_Automator.spec — Spec PyInstaller pour TCO Automator v2.2
Build : pyinstaller TCO_Automator.spec --clean
"""

import os
from PyInstaller.utils.hooks import collect_all, collect_submodules

block_cipher = None

# ---------------------------------------------------------------------------
# Donnees statiques a bundler
# ---------------------------------------------------------------------------

datas = [
    ('app.py',     '.'),
    ('config.py',  '.'),
    ('logger.py',  '.'),
    ('core',       'core'),
    ('app',        'app'),
    ('services',   'services'),
    ('odetec_logo.png', '.'),
    ('tco.png',         '.'),
]

if os.path.exists('Template_DPGF'):
    datas.append(('Template_DPGF', 'Template_DPGF'))

# ---------------------------------------------------------------------------
# Hidden imports
# ---------------------------------------------------------------------------

hiddenimports = [
    'dotenv', 'dotenv.main',
    'xlrd', 'pyxlsb',
    'openpyxl', 'openpyxl.cell._writer',
    'openpyxl.styles.fills', 'openpyxl.styles.fonts',
    'openpyxl.styles.borders', 'openpyxl.styles.numbers',
    'openpyxl.styles.alignment',
    'rapidfuzz', 'rapidfuzz.fuzz', 'rapidfuzz.process',
    'pdfplumber',
    'fitz',
    'pandas', 'pandas.core.arrays.integer',
    'pandas.core.arrays.floating', 'pandas.io.formats.excel',
    'streamlit', 'streamlit.web.cli',
    'streamlit.runtime.scriptrunner.magic_funcs',
    'streamlit.components.v1',
    'gzip', 'json', 'uuid', 'decimal', 'logging.handlers', 'signal',
    'zipfile', 're', 'copy',
]

hiddenimports += collect_submodules('streamlit')
hiddenimports += collect_submodules('openpyxl')
hiddenimports += collect_submodules('pdfplumber')

# ---------------------------------------------------------------------------
# Collect datas + binaries depuis les packages Python
# ---------------------------------------------------------------------------

binaries = []

for _pkg in ('streamlit', 'pandas', 'openpyxl', 'pdfplumber', 'fitz', 'rapidfuzz', 'xlrd', 'pyxlsb'):
    try:
        _tmp = collect_all(_pkg)
        datas    += _tmp[0]
        binaries += _tmp[1]
        hiddenimports += _tmp[2]
    except Exception:
        pass

# ---------------------------------------------------------------------------
# Analysis
# ---------------------------------------------------------------------------

a = Analysis(
    ['run_app.py'],
    pathex=['.'],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'pytest', 'ruff', 'pip', 'setuptools', 'wheel',
        'tkinter', 'matplotlib', 'scipy', 'sklearn',
        'IPython', 'notebook', 'jupyter',
    ],
    noarchive=False,
    optimize=0,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# ---------------------------------------------------------------------------
# EXE
# ---------------------------------------------------------------------------

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name='TCO_Automator',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='tco.png' if os.path.exists('tco.png') else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=True,
    upx_exclude=[],
    name='TCO_Automator',
)
