@echo off
cd /d %~dp0
set PYTHONIOENCODING=utf-8
python income_report.py
pause
