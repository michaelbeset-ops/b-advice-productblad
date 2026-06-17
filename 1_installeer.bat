@echo off
echo ================================================
echo  Productblad Generator - Installatie
echo ================================================
echo.
echo Pakketten installeren... even geduld...
echo.
py -m pip install selenium openpyxl pyproj watchdog webdriver-manager Pillow requests anthropic
echo.
echo ================================================
echo  BELANGRIJK: Claude AI instellen voor Street View
echo ================================================
echo.
echo  Voor automatische container-detectie heb je een
echo  gratis Anthropic API key nodig:
echo  1. Ga naar https://console.anthropic.com
echo  2. Maak een account en kopieer je API key
echo  3. Voer het volgende commando in PowerShell in:
echo.
echo  [System.Environment]::SetEnvironmentVariable("ANTHROPIC_API_KEY","jouw-key-hier","User")
echo.
echo  Zonder API key werkt alles nog steeds, alleen
echo  de container-detectie in Street View is dan uit.
echo ================================================
echo  Klaar! Je kunt nu 3_start_watcher.bat gebruiken.
echo ================================================
pause
