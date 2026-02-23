@echo off
title Excel Automation — Project Management
color 0B
pip install pandas openpyxl xlrd colorama tabulate numpy --quiet --upgrade
cd /d "%~dp0"
python main.py project_mgmt
pause
