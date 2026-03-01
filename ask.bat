@echo off
REM ============================================================
REM  ask.bat — Natural Language CLI Launcher
REM  Usage: ask.bat sales.xlsx "remove duplicates"
REM         ask.bat reports\ "consolidate all files"
REM         ask.bat             (interactive mode)
REM ============================================================

set PYTHON=python
python --version >nul 2>&1
if errorlevel 1 (
    set PYTHON=py
    py --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. Install from https://python.org
        pause
        exit /b 1
    )
)

%PYTHON% ask.py %*
pause
