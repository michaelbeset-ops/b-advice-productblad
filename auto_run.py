# -*- coding: utf-8 -*-
"""
Volledig automatische productblad-generator.

Volgorde:
  1. Lees data uit dict (komt uit input.json via watcher, of direct aanroepen)
  2. Open Chrome met alle 6 websites
  3. Cookies/popups automatisch wegklikken, inzoomen, zoekresultaten aanklikken
  4. Neem automatisch een screenshot van elk tabblad
  5. Genereer de Excel met data + screenshots

Gebruik vanuit terminal (voor testen):
    python auto_run.py input_zone/voorbeeld_input.json
"""

import io
import json
import logging
import os
import sys
import time

from PIL import Image

FILE_LOCATION = os.path.dirname(os.path.realpath(__file__))
DEPS = os.path.join(FILE_LOCATION, "Dependencies")

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# Crop-regio's (left, top, right, bottom) in browser-viewport pixels
# Ingesteld voor een 1920x1080 scherm. Pas aan als screenshots niet kloppen.
# ---------------------------------------------------------------------------
CROP_REGIO = {
    "Kaart":       (380,  200, 1500, 820),
    "Kadaster":    (380,  200, 1700, 900),
    "Luchtfoto":   (880,   80, 1520, 900),
    "GPS":         (160,  370, 1180, 790),
    "Locatiefoto": (280,   60, 1530, 780),
    "Loopafstand": (250,   60, 1480, 760),
}

WACHT_TIJD = {
    "Kaart":       6.0,
    "Kadaster":    5.0,
    "Luchtfoto":   5.0,
    "GPS":         4.0,
    "Locatiefoto": 6.0,
    "Loopafstand": 5.0,
}


# ---------------------------------------------------------------------------
# Hulpfuncties
# ---------------------------------------------------------------------------

def _klik_als_aanwezig(driver, by, selector, timeout=5):
    """Klik op element als het bestaat — gooit geen fout als het er niet is."""
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, selector))
        )
        el.click()
        time.sleep(0.5)
        return True
    except Exception:
        return False


def _sluit_google_cookies(driver):
    """Klik 'Alles accepteren' weg op Google-pagina's."""
    from selenium.webdriver.common.by import By
    pogingen = [
        (By.XPATH, '//button[.//span[contains(text(),"Alles accepteren")]]'),
        (By.XPATH, '//button[contains(text(),"Alles accepteren")]'),
        (By.XPATH, '//button[.//span[contains(text(),"Accept all")]]'),
        (By.XPATH, '//form[contains(@action,"consent")]//button'),
    ]
    for by, sel in pogingen:
        if _klik_als_aanwezig(driver, by, sel, timeout=4):
            log.info("Google cookiebanner weggeklikt.")
            time.sleep(1)
            return


def _sluit_algemene_cookiebanner(driver):
    """Generieke cookiebanner wegklikken (werkt op veel Nederlandse sites)."""
    from selenium.webdriver.common.by import By
    pogingen = [
        (By.XPATH, '//button[contains(translate(text(),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"accepteer")]'),
        (By.XPATH, '//button[contains(translate(text(),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"akkoord")]'),
        (By.XPATH, '//button[contains(translate(text(),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"toestaan")]'),
        (By.XPATH, '//button[contains(translate(text(),"ABCDEFGHIJKLMNOPQRSTUVWXYZ","abcdefghijklmnopqrstuvwxyz"),"accept")]'),
        (By.CSS_SELECTOR, 'button[id*="accept"], button[class*="accept"], button[id*="cookie"], button[class*="cookie"]'),
    ]
    for by, sel in pogingen:
        if _klik_als_aanwezig(driver, by, sel, timeout=3):
            log.info("Cookiebanner weggeklikt.")
            time.sleep(0.5)
            return


def _screenshot_viewport(driver) -> Image.Image:
    png_bytes = driver.get_screenshot_as_png()
    return Image.open(io.BytesIO(png_bytes))


def _crop_en_sla_op(driver, naam: str, output_dir: str):
    img = _screenshot_viewport(driver)
    regio = CROP_REGIO.get(naam)
    if regio:
        img = img.crop(regio)
    pad = os.path.join(output_dir, f"{naam}.png")
    img.save(pad)
    log.info(f"Screenshot opgeslagen: {naam}.png")
    return pad


def _zoom_kaart(driver, stappen: int):
    """Zoom in of uit op een kaart via de + / - knoppen in de URL-balk simuleren."""
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains
    body = driver.find_element("tag name", "body")
    zoom_key = Keys.ADD if stappen > 0 else Keys.SUBTRACT
    for _ in range(abs(stappen)):
        ActionChains(driver).key_down(Keys.CONTROL).send_keys(zoom_key).key_up(Keys.CONTROL).perform()
        time.sleep(0.3)


# ---------------------------------------------------------------------------
# Per-website functies
# ---------------------------------------------------------------------------

def _tab_bagviewer(driver, wacht_klaar, target_address, output_dir):
    log.info("Tab 1: BAGviewer (Kaart)")
    driver.get(f"https://bagviewer.kadaster.nl/lvbag/bag-viewer/?searchQuery={target_address}")
    wacht_klaar()
    _sluit_algemene_cookiebanner(driver)
    time.sleep(WACHT_TIJD["Kaart"])
    _crop_en_sla_op(driver, "Kaart", output_dir)


def _tab_pdok(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    log.info("Tab 2: PDOK (Kadaster)")
    nieuw_tabblad("https://app.pdok.nl/viewer/", extra=2.5)
    _sluit_algemene_cookiebanner(driver)

    # Zoeken in zoekveld
    try:
        veld = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "ggcSearchInputId"))
        )
        veld.clear()
        veld.send_keys(target_address)
        time.sleep(1.5)
        # Klik het eerste suggestieresultaat aan
        gevonden = _klik_als_aanwezig(
            driver,
            By.CSS_SELECTOR,
            ".ggc-search-results li:first-child, .suggestion-item:first-child, "
            "[class*='result']:first-child, [class*='suggestion']:first-child",
            timeout=4
        )
        if not gevonden:
            veld.send_keys(Keys.RETURN)
    except Exception as e:
        log.warning(f"PDOK zoeken: {e}")

    time.sleep(WACHT_TIJD["Kadaster"])
    _crop_en_sla_op(driver, "Kadaster", output_dir)


def _tab_google_maps(driver, wacht_klaar, nieuw_tabblad, google_url, output_dir):
    log.info("Tab 3: Google Maps satelliet (Luchtfoto)")
    nieuw_tabblad(google_url, extra=2.0)
    _sluit_google_cookies(driver)
    time.sleep(WACHT_TIJD["Luchtfoto"])
    _crop_en_sla_op(driver, "Luchtfoto", output_dir)


def _tab_gps(driver, wacht_klaar, nieuw_tabblad, coords, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    log.info("Tab 4: GPS coördinaten")
    nieuw_tabblad(
        "https://www.gpscoordinaten.nl/converteer-gps-coordinaten.php",
        extra=2.0
    )
    _sluit_algemene_cookiebanner(driver)

    try:
        veld = driver.find_element(By.ID, "a-latlong")
        veld.clear()
        veld.send_keys(coords)
        veld.send_keys(Keys.RETURN)
    except Exception as e:
        log.warning(f"GPS zoekveld: {e}")

    time.sleep(WACHT_TIJD["GPS"])
    _crop_en_sla_op(driver, "GPS", output_dir)


def _tab_streetview(driver, wacht_klaar, nieuw_tabblad, joined_coords, output_dir):
    log.info("Tab 5: Google Street View (Locatiefoto)")
    nieuw_tabblad(
        f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={joined_coords}",
        extra=2.0
    )
    _sluit_google_cookies(driver)
    time.sleep(WACHT_TIJD["Locatiefoto"])
    _crop_en_sla_op(driver, "Locatiefoto", output_dir)


def _tab_afstandmeten(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    log.info("Tab 6: Afstandmeten.nl (Loopafstand)")
    nieuw_tabblad("https://afstandmeten.nl/", extra=3.0)
    _sluit_algemene_cookiebanner(driver)

    try:
        veld = driver.find_element(By.ID, "qId")
        veld.clear()
        veld.send_keys(target_address)
        veld.send_keys(Keys.RETURN)
    except Exception as e:
        log.warning(f"Afstandmeten zoekveld: {e}")

    time.sleep(WACHT_TIJD["Loopafstand"])
    _crop_en_sla_op(driver, "Loopafstand", output_dir)


# ---------------------------------------------------------------------------
# Hoofdfunctie: alles openen en screenshotten
# ---------------------------------------------------------------------------

def open_browsers_en_screenshot(data: dict, output_dir: str):
    from selenium import webdriver
    from selenium.webdriver.chrome.service import Service
    from selenium.webdriver.support.ui import WebDriverWait

    straat = str(data.get("straat", ""))
    huisnr = str(data.get("huisnummer", ""))
    toev   = str(data.get("toevoeging", ""))
    plaats = str(data.get("plaats", ""))
    coords = str(data.get("coordinaten", ""))

    target_address = f"{straat} {huisnr}{toev}, {plaats}"
    joined_coords  = coords.replace(" ", "")
    google_url     = (
        f"https://www.google.nl/maps/place/"
        f"{straat} {huisnr}{toev}, {plaats}/data=!3m1!1e3"
    )

    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--lang=nl")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    # Accepteer automatisch geolocation (minder popups)
    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.geolocation": 1,
        "profile.default_content_setting_values.notifications": 2,
    })

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = Service(executable_path=os.path.join(DEPS, "chromedriver.exe"))

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

    try:
        _tab_bagviewer(driver, wacht_klaar, target_address, output_dir)
        _tab_pdok(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir)
        _tab_google_maps(driver, wacht_klaar, nieuw_tabblad, google_url, output_dir)
        _tab_gps(driver, wacht_klaar, nieuw_tabblad, coords, output_dir)
        _tab_streetview(driver, wacht_klaar, nieuw_tabblad, joined_coords, output_dir)
        _tab_afstandmeten(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir)
    finally:
        driver.quit()
        log.info("Browser gesloten.")


# ---------------------------------------------------------------------------
# run_volledig: screenshots + Excel in één aanroep
# ---------------------------------------------------------------------------

def run_volledig(data: dict) -> str:
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
# Directe aanroep:  python auto_run.py input_zone/voorbeeld_input.json
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

    with open(sys.argv[1], "r", encoding="utf-8") as f:
        data = json.load(f)

    resultaat = run_volledig(data)
    print(f"\nKlaar! Bestand staat op:\n{resultaat}")
