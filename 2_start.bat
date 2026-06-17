@echo off
cd /d "%~dp0"
python start.py
if %errorlevel% neq 0 (
    echo.
    echo Er is een fout opgetreden. Zie de foutmelding hierboven.
    pause
)
