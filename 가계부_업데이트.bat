@echo off
chcp 65001 >nul
cd /d "%~dp0"
python scripts\build_ledger.py
echo.
pause
