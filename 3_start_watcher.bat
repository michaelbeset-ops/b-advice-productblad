@echo off
cd /d "%~dp0"
echo ============================================
echo  B-Advice Productblad Watcher
echo ============================================
echo.
echo  Drop een input.json in de map input_zone\
echo  en het productblad wordt automatisch gemaakt.
echo.
echo  Stop met Ctrl+C
echo ============================================
echo.
py watcher.py
if %errorlevel% neq 0 (
    echo.
    echo Er is een fout opgetreden. Zie hierboven.
    pause
)
