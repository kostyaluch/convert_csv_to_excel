@echo off
REM Build and Verification Script for CSV to Excel Converter
REM This script builds the executable and performs basic verification

echo ========================================
echo CSV to Excel Converter - Build Script
echo ========================================
echo.

REM Step 1: Check Python installation
echo [1/6] Checking Python installation...
python --version
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    pause
    exit /b 1
)
echo ✓ Python found
echo.

REM Step 2: Check required files
echo [2/6] Checking required files...
if not exist "convert_csv_to_excel_v3.py" (
    echo ERROR: convert_csv_to_excel_v3.py not found
    pause
    exit /b 1
)
if not exist "convert_csv_to_excel_v3.spec" (
    echo ERROR: convert_csv_to_excel_v3.spec not found
    pause
    exit /b 1
)
if not exist "logo.ico" (
    echo ERROR: logo.ico not found
    pause
    exit /b 1
)
if not exist "header_map.json" (
    echo ERROR: header_map.json not found
    pause
    exit /b 1
)
echo ✓ All required files found
echo.

REM Step 3: Install dependencies
echo [3/6] Installing dependencies...
pip install -r requirements.txt
if errorlevel 1 (
    echo ERROR: Failed to install dependencies
    pause
    exit /b 1
)
echo ✓ Dependencies installed
echo.

REM Step 4: Clean previous build
echo [4/6] Cleaning previous build...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
echo ✓ Build directories cleaned
echo.

REM Step 5: Build executable
echo [5/6] Building executable with PyInstaller...
echo This may take several minutes...
pyinstaller convert_csv_to_excel_v3.spec
if errorlevel 1 (
    echo ERROR: Build failed
    pause
    exit /b 1
)
echo ✓ Build completed
echo.

REM Step 6: Verify output
echo [6/6] Verifying build output...
if not exist "dist\CSVtoExcel.exe" (
    echo ERROR: CSVtoExcel.exe was not created
    pause
    exit /b 1
)
echo ✓ CSVtoExcel.exe created successfully
echo.

REM Get file size
for %%A in ("dist\CSVtoExcel.exe") do set size=%%~zA
set /a sizeMB=%size% / 1048576
echo File size: %sizeMB% MB
echo.

echo ========================================
echo BUILD SUCCESSFUL!
echo ========================================
echo.
echo Executable location: dist\CSVtoExcel.exe
echo.
echo Next steps:
echo 1. Test the executable by running: dist\CSVtoExcel.exe
echo 2. Verify all features work correctly
echo 3. Copy CSVtoExcel.exe to your distribution folder
echo.
pause
