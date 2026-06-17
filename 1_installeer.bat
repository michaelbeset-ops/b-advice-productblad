@echo off
echo ================================================
echo  Productblad Generator - Installatie
echo ================================================
echo.
echo Pakketten installeren... even geduld...
echo.
py -m pip install selenium openpyxl pyproj watchdog webdriver-manager Pillow requests
echo.
echo ================================================
echo  Klaar! Je kunt nu 2_start_watcher.bat gebruiken.
echo ================================================
pause
