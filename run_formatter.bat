@echo off
title Excel Automation — Formatter Module
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  — FORMATTER MODULE
echo    Charts, Traffic Lights, Color Scale, Table Styles, Totals
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

echo.
echo  Launching Formatter module...
echo  ============================================================
echo.

cd /d "%~dp0"
python main.py format

echo.
echo  ============================================================
echo   Session ended. Check the 'output' folder for your files.
echo  ============================================================
echo.
pause
