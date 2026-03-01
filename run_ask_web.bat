@echo off
REM ============================================================
REM  run_ask_web.bat — Streamlit Web UI Launcher
REM  Opens browser UI at http://localhost:8501
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

echo [INFO] Installing/checking streamlit...
%PYTHON% -m pip install streamlit --quiet

echo.
echo [INFO] Starting Excel NL Toolkit...
echo [INFO] Browser will open at: http://localhost:8501
echo [INFO] Press Ctrl+C in this window to stop.
echo.

%PYTHON% -m streamlit run ask_web.py
pause
