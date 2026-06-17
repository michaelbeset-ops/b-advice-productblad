# -*- coding: utf-8 -*-
"""
Volledig automatische productblad-generator.

Volgorde:
  1. Lees data uit dict (komt uit input.json via watcher, of direct aanroepen)
  2. Open Chrome met alle 6 websites
  3. Neem automatisch een screenshot van elk tabblad
  4. Genereer de Excel met data + screenshots
  5. Geef pad naar het klaarstaande .xlsm bestand terug

Gebruik vanuit Python:
    from auto_run import run_volledig
    pad = run_volledig(data_dict)

Gebruik vanuit terminal (voor testen):
    python auto_run.py input_zone/voorbeeld_input.json
"""

import io
import json
import logging
import os
import sys
import time

from PIL import Image  # pip install Pillow

FILE_LOCATION = os.path.dirname(os.path.realpath(__file__))
DEPS = os.path.join(FILE_LOCATION, "Dependencies")

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Crop-regio's per tabblad (linker-x, boven-y, rechter-x, onder-y)
# gemeten in browser-viewport pixels op een 1920x1080 scherm.
# Pas deze aan als jouw scherm een andere resolutie heeft.
# ---------------------------------------------------------------------------
CROP_REGIO = {
    "Kaart":       (400,  240, 1495, 805),   # BAGviewer
    "Kadaster":    (400,  240, 1700, 895),   # PDOK
    "Luchtfoto":   (910,  100, 1510, 890),   # Google Maps satelliet
    "GPS":         (175,  390, 1175, 777),   # gpscoordinaten.nl
    "Locatiefoto": (300,   40, 1520, 760),   # Google Street View
    "Loopafstand": (270,   40, 1470, 740),   # afstandmeten.nl
}

# Wachttijden per website (seconden) zodat kaarten volledig laden
WACHT_TIJD = {
    "Kaart":       5.0,
    "Kadaster":    4.0,
    "Luchtfoto":   4.0,
    "GPS":         3.0,
    "Locatiefoto": 5.0,
    "Loopafstand": 4.0,
}


def _screenshot_viewport(driver) -> Image.Image:
    """Maak een screenshot van de volledige browser-viewport als PIL Image."""
    png_bytes = driver.get_screenshot_as_png()
    return Image.open(io.BytesIO(png_bytes))


def _crop_en_sla_op(driver, naam: str, output_dir: str):
    """Screenshot maken en bijsnijden naar de geconfigureerde regio."""
    img = _screenshot_viewport(driver)
    regio = CROP_REGIO.get(naam)
    if regio:
        img = img.crop(regio)
    pad = os.path.join(output_dir, f"{naam}.png")
    img.save(pad)
    log.info(f"Screenshot opgeslagen: {naam}.png")
    return pad


def open_browsers_en_screenshot(data: dict, output_dir: str):
    """
    Open Chrome, bezoek alle 6 websites en neem screenshots.
    Alle screenshots worden opgeslagen in output_dir.
    """
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait

    straat   = str(data.get("straat", ""))
    huisnr   = str(data.get("huisnummer", ""))
    toev     = str(data.get("toevoeging", ""))
    plaats   = str(data.get("plaats", ""))
    coords   = str(data.get("coordinaten", ""))

    target_address = f"{straat} {huisnr}{toev}, {plaats}"
    joined_coords  = coords.replace(" ", "")
    google_url     = (
        f"https://www.google.nl/maps/place/"
        f"{straat} {huisnr}{toev}, {plaats}/data=!3m1!1e3"
    )

    # Chrome opties — verwijder '--headless' als je de browser wilt zien
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    # ChromeDriver automatisch ophalen
    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        chrome_driver = os.path.join(DEPS, "chromedriver.exe")
        service = Service(executable_path=chrome_driver)

    driver = webdriver.Chrome(service=service, options=options)

    def wacht_klaar(timeout=20):
        try:
            WebDriverWait(driver, timeout).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
        except Exception:
            pass

    def nieuw_tabblad(url, extra=0.0):
        driver.execute_script("window.open();")
        driver.switch_to.window(driver.window_handles[-1])
        driver.get(url)
        wacht_klaar()
        if extra > 0:
            time.sleep(extra)

    def zoek_in_veld(element_id, tekst):
        try:
            veld = driver.find_element(By.ID, element_id)
            veld.clear()
            veld.send_keys(tekst)
            veld.send_keys(Keys.RETURN)
        except Exception as e:
            log.warning(f"Zoekveld '{element_id}' niet gevonden: {e}")

    try:
        log.info("Browser openen — Tab 1: BAGviewer (Kaart)")
        driver.get(f"https://bagviewer.kadaster.nl/lvbag/bag-viewer/?searchQuery={target_address}")
        wacht_klaar()
        time.sleep(WACHT_TIJD["Kaart"])
        _crop_en_sla_op(driver, "Kaart", output_dir)

        log.info("Tab 2: PDOK (Kadaster)")
        nieuw_tabblad("https://app.pdok.nl/viewer/", extra=2.0)
        zoek_in_veld("ggcSearchInputId", target_address)
        time.sleep(WACHT_TIJD["Kadaster"])
        _crop_en_sla_op(driver, "Kadaster", output_dir)

        log.info("Tab 3: Google Maps satelliet (Luchtfoto)")
        nieuw_tabblad(google_url, extra=WACHT_TIJD["Luchtfoto"])
        _crop_en_sla_op(driver, "Luchtfoto", output_dir)

        log.info("Tab 4: GPS coördinaten")
        nieuw_tabblad(
            "https://www.gpscoordinaten.nl/converteer-gps-coordinaten.php",
            extra=2.0
        )
        zoek_in_veld("a-latlong", coords)
        time.sleep(WACHT_TIJD["GPS"])
        _crop_en_sla_op(driver, "GPS", output_dir)

        log.info("Tab 5: Google Street View (Locatiefoto)")
        nieuw_tabblad(
            f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={joined_coords}",
            extra=WACHT_TIJD["Locatiefoto"]
        )
        _crop_en_sla_op(driver, "Locatiefoto", output_dir)

        log.info("Tab 6: Afstandmeten.nl (Loopafstand)")
        nieuw_tabblad("https://afstandmeten.nl/", extra=3.0)
        zoek_in_veld("qId", target_address)
        time.sleep(WACHT_TIJD["Loopafstand"])
        _crop_en_sla_op(driver, "Loopafstand", output_dir)

    finally:
        driver.quit()
        log.info("Browser gesloten.")


def run_volledig(data: dict) -> str:
    """
    Hoofdfunctie: screenshots + Excel in één aanroep.
    Geeft het pad naar het klaarstaande .xlsm-bestand terug.
    """
    import re
    from productblad_core import generate_excel

    straat   = str(data.get("straat", ""))
    huisnr   = str(data.get("huisnummer", ""))
    postcode = str(data.get("postcode", ""))
    plaats   = str(data.get("plaats", ""))

    letters = re.findall(r'\D+', postcode)
    pc_letters = letters[0].strip() if letters else postcode
    location_code    = f"{pc_letters}{huisnr} - {straat}"
    veilige_code     = re.sub(r'[\\/*?:"<>|]', '_', location_code).strip()
    veilige_gemeente = re.sub(r'[\\/*?:"<>|]', '_', plaats).strip()
    output_dir = os.path.join(FILE_LOCATION, "Pythonwerk", veilige_gemeente, veilige_code)
    os.makedirs(output_dir, exist_ok=True)

    log.info("=== Stap 1: Screenshots maken ===")
    open_browsers_en_screenshot(data, output_dir)

    log.info("=== Stap 2: Excel genereren ===")
    excel_pad = generate_excel(data)

    log.info(f"=== Klaar! Excel: {excel_pad} ===")
    return excel_pad


# ---------------------------------------------------------------------------
# Directe aanroep vanuit terminal:  python auto_run.py mijn_input.json
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%H:%M:%S",
    )

    if len(sys.argv) < 2:
        print("Gebruik: python auto_run.py <pad_naar_input.json>")
        sys.exit(1)

    json_pad = sys.argv[1]
    with open(json_pad, "r", encoding="utf-8") as f:
        data = json.load(f)

    resultaat = run_volledig(data)
    print(f"\nKlaar! Bestand staat op:\n{resultaat}")
