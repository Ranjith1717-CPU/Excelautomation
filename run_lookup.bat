@echo off
title Excel Automation — Lookup Module
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  — LOOKUP MODULE
echo    VLOOKUP, Fuzzy Match, Multi-Key Join, Reverse Lookup, Enrich
echo  ============================================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python not found on PATH. Install Python 3.8+ first.
    pause
    exit /b 1
)

echo  Installing / updating required packages...
pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade
echo  Installing optional fuzzy matching (rapidfuzz)...
pip install rapidfuzz --quiet

echo.
echo  Launching Lookup module...
echo  ============================================================
echo.

cd /d "%~dp0"
python main.py lookup

echo.
echo  ============================================================
echo   Session ended. Check the 'output' folder for your files.
echo  ============================================================
echo.
pause
