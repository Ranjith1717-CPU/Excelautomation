@echo off
title Excel Automation — Inventory
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  v2.0  ^|  Inventory
echo  ============================================================
echo.

:: ── Detect Python ───────────────────────────────────────────────────────────
set PYTHON=
python --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=python
    goto :found_python
)
py --version >nul 2>&1
if %errorlevel% equ 0 (
    set PYTHON=py
    goto :found_python
)
echo  [ERROR] Python not found on PATH.
echo  Install Python 3.8+ from https://www.python.org/downloads/
echo  Make sure to tick "Add Python to PATH" during setup.
pause
exit /b 1

:found_python
for /f "tokens=*" %%v in ('%PYTHON% --version 2^>^&1') do echo  Found: %%v

:: ── Install dependencies ─────────────────────────────────────────────────────
echo.
echo  Installing / updating required packages...
%PYTHON% -m pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade
if %errorlevel% neq 0 (
    echo  [WARNING] Some packages may not have installed. Trying anyway...
) else (
    echo  Packages ready.
)

:: ── Launch ───────────────────────────────────────────────────────────────────
echo.
cd /d "%~dp0"
%PYTHON% cli.py

echo.
echo  ============================================================
echo   Done. Output files are in the 'output' subfolder.
echo  ============================================================
echo.
pause
