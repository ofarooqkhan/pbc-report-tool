@echo off
REM ═══════════════════════════════════════════════════════════════
REM PBC Report Generator v3.0 — Windows Build Script
REM Builds a standalone .exe using PyInstaller
REM ═══════════════════════════════════════════════════════════════

echo PBC Report Generator — Windows Build
echo =====================================

REM Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python not found. Install Python 3.9+ first.
    pause
    exit /b 1
)

REM Install dependencies
echo Installing dependencies...
pip install --upgrade pip
pip install pandas openpyxl numpy xlrd pyinstaller

REM Clean previous builds
echo Cleaning previous builds...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

REM Build
echo Building Windows .exe...
pyinstaller ^
    --name "PBC Report Generator" ^
    --windowed ^
    --onefile ^
    --noconfirm ^
    --clean ^
    --hidden-import pandas ^
    --hidden-import openpyxl ^
    --hidden-import numpy ^
    --hidden-import xlrd ^
    --collect-submodules openpyxl ^
    pbc_report_tool_v3.py

echo.
echo =======================================
echo BUILD COMPLETE!
echo =======================================
echo.
echo Windows exe: dist\PBC Report Generator.exe
echo.
echo To run: double-click "dist\PBC Report Generator.exe"
echo.
pause
