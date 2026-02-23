@echo off
title Excel Automation — Clean Data
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  — CLEAN DATA
echo    Duplicates, Blanks, Dates, Types, Outliers, Full Clean
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
echo  Launching Clean module...
echo  ============================================================
echo.

cd /d "%~dp0"
python main.py clean

echo.
echo  ============================================================
echo   Session ended. Check the 'output' folder for your files.
echo  ============================================================
echo.
pause
