@echo off
title Excel Automation — Lookup
color 0B
pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade

pip install rapidfuzz --quiet
cd /d "%~dp0"
python cli.py
pause
