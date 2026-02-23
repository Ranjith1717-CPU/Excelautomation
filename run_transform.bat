@echo off
title Excel Automation — Transform Data
color 0B

echo.
echo  ============================================================
echo    EXCEL AUTOMATION TOOLKIT  — TRANSFORM DATA
echo    Pivot, Unpivot, Transpose, Split, Reshape, Rank
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
echo  Launching Transform module...
echo  ============================================================
echo.

cd /d "%~dp0"
python main.py transform

echo.
echo  ============================================================
echo   Session ended. Check the 'output' folder for your files.
echo  ============================================================
echo.
pause
