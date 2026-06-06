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
        ('header_map.json', '.'),  # Include default header map
        ('logo.ico', '.'),  # Include logo for runtime access
    ],
    hiddenimports=[
        # Core libraries
        'pandas', 
        'openpyxl', 
        'openpyxl.styles', 
        'openpyxl.styles.fonts',
        'openpyxl.styles.alignment',
        'openpyxl.utils', 
        'openpyxl.cell',
        'openpyxl.cell.cell',
        'openpyxl.worksheet',
        'openpyxl.worksheet.worksheet',
        'openpyxl.workbook',
        'openpyxl.workbook.workbook',
        'openpyxl.reader.excel',
        'openpyxl.writer.excel',
        # Pandas dependencies
        'pandas._libs',
        'pandas._libs.tslibs',
        'pandas._libs.tslibs.timedeltas',
        'pandas._libs.tslibs.np_datetime',
        'pandas._libs.tslibs.nattype',
        'pandas._libs.tslibs.timestamps',
        'pandas._libs.skiplist',
        'pandas.core.arrays.datetimes',
        'pytz',
        'pytz.tzfile',
        'tzdata',
        'dateutil',
        'dateutil.parser',
        'dateutil.tz',
        # XML processing for openpyxl
        'et_xmlfile',
        'et_xmlfile.xmlfile',
        'xml.etree.ElementTree',
        'lxml',
        'lxml.etree',
        # Standard library
        'html', 
        'html.parser',
        're', 
        'queue', 
        'threading', 
        'json', 
        'tkinterdnd2',
        'tkinter',
        'tkinter.filedialog',
        'tkinter.messagebox',
        'tkinter.scrolledtext',
        'tkinter.ttk',
        # Additional utilities
        'encodings',
        'encodings.utf_8',
        'encodings.cp1251',
        'encodings.cp1252',
    ],
    hookspath=['hooks'],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',  # Exclude unnecessary packages to reduce size
        'scipy',
        'numpy.testing',
        'pytest',
        'IPython',
        'notebook',
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
    # Note: version parameter is not supported in EXE() for one-file mode
    # To add version info, use a custom version file with --version-file flag
)
