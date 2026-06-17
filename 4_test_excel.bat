@echo off
cd /d "%~dp0"
echo ============================================
echo  Test: alleen Excel genereren (geen browser)
echo ============================================
echo.
py -c "
import json, sys, traceback

print('Stap 1: JSON inlezen...')
try:
    with open('input_zone/voorbeeld_input.json', encoding='utf-8') as f:
        data = json.load(f)
    print('  OK')
except Exception as e:
    print('  FOUT:', e)
    sys.exit(1)

print('Stap 2: openpyxl importeren...')
try:
    from openpyxl import load_workbook
    print('  OK')
except Exception as e:
    print('  FOUT - voer 1_installeer.bat opnieuw uit:', e)
    sys.exit(1)

print('Stap 3: pyproj importeren...')
try:
    from pyproj import Transformer
    print('  OK')
except Exception as e:
    print('  FOUT - voer 1_installeer.bat opnieuw uit:', e)
    sys.exit(1)

print('Stap 4: Excel template vinden...')
import os
template = os.path.join(os.path.dirname(os.path.abspath('.')), 'Dependencies', 'PythonWerkProjectblad.xlsm')
template2 = os.path.join(os.getcwd(), 'Dependencies', 'PythonWerkProjectblad.xlsm')
if os.path.exists(template2):
    print('  OK:', template2)
else:
    print('  FOUT: bestand niet gevonden op:', template2)
    sys.exit(1)

print('Stap 5: Excel genereren...')
try:
    from productblad_core import generate_excel
    pad = generate_excel(data)
    print('  OK! Excel staat op:', pad)
except Exception as e:
    print('  FOUT:')
    traceback.print_exc()
    sys.exit(1)

print()
print('Alles geslaagd!')
"
echo.
if %errorlevel% neq 0 (
    echo Er is een fout opgetreden. Zie hierboven welke stap mislukt is.
)
pause
