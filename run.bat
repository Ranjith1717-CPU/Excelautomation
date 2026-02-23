@echo off
title Excel Automation Toolkit v1.0
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  v1.0
echo    Powered by Python + pandas
echo  ============================================================
echo.

:: ── Check Python ────────────────────────────────────────────────────────────
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] Python not found on PATH!
    echo.
    echo  Please install Python 3.8 or later from:
    echo     https://www.python.org/downloads/
    echo.
    echo  Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

for /f "tokens=*" %%v in ('python --version 2^>^&1') do (
    echo  Python found: %%v
)

:: ── Check pip ────────────────────────────────────────────────────────────────
pip --version >nul 2>&1
if %errorlevel% neq 0 (
    echo  [ERROR] pip not found. Please re-install Python with pip enabled.
    pause
    exit /b 1
)

:: ── Install / Update dependencies ───────────────────────────────────────────
echo.
echo  Checking and installing required packages...
echo.

pip install pandas openpyxl xlrd colorama tabulate matplotlib numpy --quiet --upgrade

if %errorlevel% neq 0 (
    echo  [WARNING] Some packages may not have installed correctly.
    echo  Attempting to continue anyway...
    echo.
) else (
    echo  All packages are ready.
)

:: ── Run the toolkit ──────────────────────────────────────────────────────────
echo.
echo  Starting Excel Automation Toolkit...
echo  ============================================================
echo.

cd /d "%~dp0"
python main.py

:: ── On exit ──────────────────────────────────────────────────────────────────
echo.
echo  ============================================================
echo   Session ended. Check the 'output' folder for your files.
echo  ============================================================
echo.
pause
