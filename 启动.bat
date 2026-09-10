@echo off
chcp 65001 >nul 2>&1
cd /d "<PROJECT_ROOT>"
python main.py
pause
