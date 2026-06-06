# -*- mode: python ; coding: utf-8 -*-

import sys
from pathlib import Path

# Get tkinterdnd2 package path
import tkinterdnd2
tkdnd_path = Path(tkinterdnd2.__file__).parent

a = Analysis(
    ['convert_csv_to_excel_v3.py'],
    pathex=[],
    binaries=[],
    datas=[(str(tkdnd_path / 'tkdnd'), 'tkdnd')],
    hiddenimports=['pandas', 'openpyxl', 'openpyxl.styles', 'openpyxl.utils', 'html', 're', 'queue', 'threading', 'json', 'tkinterdnd2'],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='convert_csv_to_excel_v3',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=['logo.ico'],
)
