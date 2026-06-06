# PyInstaller hook for openpyxl
# This ensures all necessary openpyxl submodules are included

from PyInstaller.utils.hooks import collect_submodules, collect_data_files

hiddenimports = collect_submodules('openpyxl')
datas = collect_data_files('openpyxl', include_py_files=True)
