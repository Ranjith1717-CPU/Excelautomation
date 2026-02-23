@echo off
title Excel Automation — Sales
color 0B
pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade
cd /d "%~dp0"
python cli.py
pause
