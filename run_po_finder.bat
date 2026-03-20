@echo off
title PO XML Finder - Launcher

:: ── Check Python ─────────────────────────────────────────────────────────────
python --version >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python is not installed or not in PATH.
    echo Please download it from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: ── Install dependencies if missing ──────────────────────────────────────────
echo Checking dependencies...
python -c "import pandas" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Installing pandas...
    python -m pip install pandas openpyxl --quiet
)

python -c "import openpyxl" >nul 2>&1
IF %ERRORLEVEL% NEQ 0 (
    echo Installing openpyxl...
    python -m pip install openpyxl --quiet
)

:: ── Launch app ────────────────────────────────────────────────────────────────
echo Launching PO XML Finder...
start "" pythonw "%~dp0po_finder.py"