@echo off
cd /d %~dp0
set PYTHONIOENCODING=utf-8
python item_sync.py
pause
