# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller spec file for CSV to Excel Converter v3
This configuration creates a standalone executable with:
- Embedded logo icon
- No console window
- UPX compression
- All required dependencies bundled
"""

import sys
from pathlib import Path

# Get tkinterdnd2 package path for drag-and-drop support
import tkinterdnd2
tkdnd_path = Path(tkinterdnd2.__file__).parent

# Version information (for Windows executable properties)
version_info = (
    'VSVersionInfo',
    'StringFileInfo',
    [
        ('040904E4', {
            'CompanyName': 'CSV to Excel Tools',
            'FileDescription': 'CSV to Excel Converter - Professional batch conversion tool',
            'FileVersion': '3.0.0',
            'InternalName': 'CSVtoExcel',
            'LegalCopyright': '© 2024',
            'OriginalFilename': 'convert_csv_to_excel_v3.exe',
            'ProductName': 'CSV to Excel Converter',
            'ProductVersion': '3.0.0'
        })
    ]
)

a = Analysis(
    ['convert_csv_to_excel_v3.py'],
    pathex=[],
    binaries=[],
    datas=[
        (str(tkdnd_path / 'tkdnd'), 'tkdnd'),  # tkinterdnd2 native libraries
    ],
    hiddenimports=[
        'pandas', 
        'openpyxl', 
        'openpyxl.styles', 
        'openpyxl.utils', 
        'openpyxl.cell',
        'openpyxl.worksheet',
        'html', 
        're', 
        'queue', 
        'threading', 
        'json', 
        'tkinterdnd2'
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',  # Exclude unnecessary packages to reduce size
        'scipy',
        'numpy.testing',
        'pytest',
    ],
    noarchive=False,
    optimize=2,  # Maximum optimization
)

pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name='CSVtoExcel',  # Professional executable name
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,  # Enable UPX compression
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # No console window - professional GUI application
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='logo.ico',  # Application icon
    version='version_info',  # Windows version information
)
