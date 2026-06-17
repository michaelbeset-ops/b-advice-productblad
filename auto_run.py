# -*- coding: utf-8 -*-
"""
Volledig automatische productblad-generator met site-specifieke logica per website
en Claude AI Vision voor Street View container-detectie.
"""

import base64
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

CROP_REGIO = {
    "Kaart":       (0,    0, 1920, 1080),  # Volledig scherm, sidebar al gesloten
    "Kadaster":    (0,    0, 1920, 1080),
    "Luchtfoto":   (0,    0, 1920, 1080),
    "GPS":         (0,  300, 1200,  850),  # Lat t/m resultaat sectie
    "Locatiefoto": (0,    0, 1920, 1080),
    "Loopafstand": (0,    0, 1920, 1080),
}


# ---------------------------------------------------------------------------
# Hulpfuncties
# ---------------------------------------------------------------------------

def _wacht_op(driver, by, selector, timeout=10):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        return WebDriverWait(driver, timeout).until(
            EC.presence_of_element_located((by, selector))
        )
    except Exception:
        return None


def _klik_als_aanwezig(driver, by, selector, timeout=5):
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC
    try:
        el = WebDriverWait(driver, timeout).until(
            EC.element_to_be_clickable((by, selector))
        )
        el.click()
        time.sleep(0.6)
        return True
    except Exception:
        return False


def _screenshot_viewport(driver) -> Image.Image:
    return Image.open(io.BytesIO(driver.get_screenshot_as_png()))


def _crop_en_sla_op(driver, naam: str, output_dir: str, img: Image.Image = None):
    if img is None:
        img = _screenshot_viewport(driver)

    # Sla altijd eerst het volledige scherm op als debug-referentie
    debug_pad = os.path.join(output_dir, f"DEBUG_{naam}_volledig.png")
    img.save(debug_pad)

    regio = CROP_REGIO.get(naam)
    if regio:
        img = img.crop(regio)
    pad = os.path.join(output_dir, f"{naam}.png")
    img.save(pad)
    log.info(f"Screenshot opgeslagen: {naam}.png  (debug: DEBUG_{naam}_volledig.png)")
    return pad


def _zoom_via_knop(driver, by, selector, stappen: int):
    for _ in range(abs(stappen)):
        _klik_als_aanwezig(driver, by, selector, timeout=3)
        time.sleep(0.5)


# ---------------------------------------------------------------------------
# 1. BAGviewer
# ---------------------------------------------------------------------------

def _tab_bagviewer(driver, wacht_klaar, target_address, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    log.info("Tab 1: BAGviewer")
    driver.get("https://bagviewer.kadaster.nl/lvbag/bag-viewer/")
    wacht_klaar()
    time.sleep(3)

    # Adres invoeren in zoekveld
    zoekveld_selectors = [
        (By.CSS_SELECTOR, "input[placeholder*='adres'], input[placeholder*='zoek'], input[type='search']"),
        (By.CSS_SELECTOR, "input.search-input, input#search, input[name='search']"),
        (By.XPATH, "//input[contains(@class,'search') or contains(@placeholder,'zoek') or contains(@placeholder,'adres')]"),
    ]
    zoekveld = None
    for by, sel in zoekveld_selectors:
        zoekveld = _wacht_op(driver, by, sel, timeout=5)
        if zoekveld:
            break

    if zoekveld:
        zoekveld.clear()
        zoekveld.send_keys(target_address)
        time.sleep(1.5)

        # Klik op eerste zoeksuggestie of druk op zoekknop
        gevonden = _klik_als_aanwezig(
            driver, By.CSS_SELECTOR,
            ".search-suggestion:first-child, .autocomplete-item:first-child, "
            "li.suggestion:first-child, [class*='suggestion']:first-child, "
            "[class*='result']:first-child, ul li:first-child",
            timeout=4
        )
        if not gevonden:
            # Probeer zoekknop
            _klik_als_aanwezig(driver, By.CSS_SELECTOR,
                "button[type='submit'], button.search-button, button[aria-label*='oek']",
                timeout=3)
            if not gevonden:
                zoekveld.send_keys(Keys.RETURN)

    time.sleep(4)  # Wacht op kaart laden

    # 1 stap uitzoomen
    _zoom_via_knop(driver, By.CSS_SELECTOR,
        "button[title*='uit'], button[aria-label*='uit'], "
        "button.zoom-out, a.leaflet-control-zoom-out, "
        ".ol-zoom-out, button[title='-']",
        stappen=1
    )
    time.sleep(1)

    # Sluit zijbalk (pijltje)
    zijbalk_selectors = [
        (By.CSS_SELECTOR, "button.sidebar-toggle, button[aria-label*='sluit'], "
                          "button[aria-label*='verberg'], .panel-toggle, "
                          "[class*='collapse'], [class*='close-panel']"),
        (By.XPATH, "//button[contains(@class,'toggle') or contains(@title,'Sluit') "
                   "or contains(@aria-label,'Sluit') or contains(@title,'sluit')]"),
    ]
    for by, sel in zijbalk_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=3):
            log.info("BAGviewer: zijbalk gesloten.")
            break
    time.sleep(1.5)

    _crop_en_sla_op(driver, "Kaart", output_dir)


# ---------------------------------------------------------------------------
# 2. PDOK
# ---------------------------------------------------------------------------

def _tab_pdok(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.support.ui import WebDriverWait
    from selenium.webdriver.support import expected_conditions as EC

    log.info("Tab 2: PDOK Kadastrale kaart")
    nieuw_tabblad("https://app.pdok.nl/viewer/", extra=3.0)

    # Zoektruc: verwijder laatste letter en typ opnieuw (activeert autocomplete)
    try:
        veld = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.ID, "ggcSearchInputId"))
        )
        veld.clear()
        veld.send_keys(target_address[:-1])  # alles behalve laatste letter
        time.sleep(0.5)
        veld.send_keys(target_address[-1])   # laatste letter opnieuw typen
        time.sleep(2)

        # Klik eerste suggestie
        gevonden = _klik_als_aanwezig(
            driver, By.CSS_SELECTOR,
            ".ggc-search-results li:first-child, "
            ".search-results li:first-child, "
            "[class*='result'] li:first-child, "
            "[class*='suggestion']:first-child",
            timeout=5
        )
        if not gevonden:
            veld.send_keys(Keys.RETURN)
    except Exception as e:
        log.warning(f"PDOK zoeken mislukt: {e}")

    time.sleep(4)

    # Open lagen-menu linksboven
    lagen_selectors = [
        (By.CSS_SELECTOR, "button[title*='agen'], button[aria-label*='agen'], "
                          ".layer-button, button[title*='Layer'], [class*='layer-control']"),
        (By.XPATH, "//button[contains(@title,'Laag') or contains(@title,'laag') "
                   "or contains(@aria-label,'laag') or contains(@title,'Layer')]"),
    ]
    for by, sel in lagen_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=4):
            log.info("PDOK: lagen-menu geopend.")
            time.sleep(1)
            break

    # Zet "Kadastrale kaart (WMS)" aan — zoek op tekst
    kadas_selectors = [
        (By.XPATH, "//*[contains(text(),'Kadastrale kaart') and (contains(text(),'WMS') or contains(text(),'10'))]"),
        (By.XPATH, "//*[contains(text(),'kadastralekaart') or contains(text(),'Kadastralekaart')]"),
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'kadastrale kaart')]"),
    ]
    for by, sel in kadas_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=4):
            log.info("PDOK: Kadastrale kaart laag geopend.")
            time.sleep(1)
            break

    # Zet sublagen aan: "kadastralekaart v5"
    for tekst in ["kadastralekaart v5", "KadastraleKaart v5", "kadastrale kaart v5"]:
        if _klik_als_aanwezig(driver, By.XPATH,
                f"//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'{tekst.lower()}')]",
                timeout=3):
            log.info(f"PDOK: '{tekst}' aangezet.")
            time.sleep(0.5)
            break

    # Zet sublaag "bebouwing" aan
    if _klik_als_aanwezig(driver, By.XPATH,
            "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'bebouwing')]",
            timeout=3):
        log.info("PDOK: 'bebouwing' laag aangezet.")
        time.sleep(0.5)

    # Sluit lagen-menu
    for by, sel in lagen_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=3):
            log.info("PDOK: lagen-menu gesloten.")
            break

    time.sleep(3)
    _crop_en_sla_op(driver, "Kadaster", output_dir)


# ---------------------------------------------------------------------------
# 3. Google Maps satelliet
# ---------------------------------------------------------------------------

def _tab_google_maps(driver, wacht_klaar, nieuw_tabblad, google_url, output_dir):
    from selenium.webdriver.common.by import By

    log.info("Tab 3: Google Maps satelliet")
    nieuw_tabblad(google_url, extra=2.5)

    # Cookies accepteren
    for sel in [
        '//button[.//span[contains(text(),"Alles accepteren")]]',
        '//button[contains(text(),"Alles accepteren")]',
        '//button[.//span[contains(text(),"Accept all")]]',
        '//form[contains(@action,"consent")]//button[last()]',
    ]:
        if _klik_als_aanwezig(driver, By.XPATH, sel, timeout=5):
            log.info("Google Maps: cookies geaccepteerd.")
            time.sleep(2)
            break

    # Sluit linker adrespaneel (pijltje)
    paneel_selectors = [
        (By.CSS_SELECTOR, "button[jsaction*='pane.close'], button[aria-label*='Sluiten'], "
                          "button[aria-label*='sluit'], button[data-value='Sluiten']"),
        (By.XPATH, "//button[@aria-label='Sluiten' or @aria-label='Close' or "
                   "contains(@jsaction,'back') or contains(@aria-label,'Terug')]"),
    ]
    for by, sel in paneel_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=4):
            log.info("Google Maps: adrespaneel gesloten.")
            time.sleep(1)
            break

    # 1 stap uitzoomen
    if _klik_als_aanwezig(driver, By.CSS_SELECTOR,
            "button[aria-label*='uitzoomen'], button[aria-label*='Zoom out'], "
            "button[title*='uitzoomen'], div[title*='Zoom out'] button",
            timeout=3):
        log.info("Google Maps: uitgezoomd.")
    time.sleep(3)

    _crop_en_sla_op(driver, "Luchtfoto", output_dir)


# ---------------------------------------------------------------------------
# 4. GPS coördinaten
# ---------------------------------------------------------------------------

def _tab_gps(driver, wacht_klaar, nieuw_tabblad, coords, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys

    log.info("Tab 4: GPS coördinaten")
    nieuw_tabblad(
        "https://www.gpscoordinaten.nl/converteer-gps-coordinaten.php",
        extra=2.5
    )

    # Toestemming / cookiebanner wegklikken
    for sel in [
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'toestemming')]"),
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'accepteer')]"),
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'akkoord')]"),
        (By.CSS_SELECTOR, "button[id*='accept'], button[class*='accept'], .cc-btn, .consent-btn"),
    ]:
        if _klik_als_aanwezig(driver, sel[0], sel[1], timeout=3):
            log.info("GPS: toestemming/cookie weggeklikt.")
            time.sleep(0.5)
            break

    # Coördinaten invullen
    try:
        veld = driver.find_element(By.ID, "a-latlong")
        veld.clear()
        veld.send_keys(coords)
    except Exception:
        log.warning("GPS: coördinatenveld niet gevonden.")

    # Klik op "Converteer" knop
    converteer_selectors = [
        (By.XPATH, "//button[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'converteer')]"),
        (By.XPATH, "//input[@type='submit' and contains(translate(@value,'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'converteer')]"),
        (By.CSS_SELECTOR, "button[type='submit'], input[type='submit']"),
    ]
    for by, sel in converteer_selectors:
        if _klik_als_aanwezig(driver, by, sel, timeout=4):
            log.info("GPS: Converteer geklikt.")
            break

    time.sleep(3)
    _crop_en_sla_op(driver, "GPS", output_dir)


# ---------------------------------------------------------------------------
# 5. Google Street View — met Claude AI container-detectie
# ---------------------------------------------------------------------------

def _ai_ziet_container(image_path: str) -> bool:
    """Vraag Claude of er een (ondergrondse) afvalcontainer zichtbaar is."""
    api_key = os.environ.get("ANTHROPIC_API_KEY", "")
    if not api_key:
        log.warning("Geen ANTHROPIC_API_KEY — AI container-check overgeslagen.")
        return True  # Geen check = neem eerste screenshot

    try:
        import anthropic
        client = anthropic.Anthropic(api_key=api_key)

        with open(image_path, "rb") as f:
            image_data = base64.b64encode(f.read()).decode("utf-8")

        response = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=50,
            messages=[{
                "role": "user",
                "content": [
                    {
                        "type": "image",
                        "source": {
                            "type": "base64",
                            "media_type": "image/png",
                            "data": image_data,
                        }
                    },
                    {
                        "type": "text",
                        "text": (
                            "Is er een ondergrondse afvalcontainer, inzamelput of containerput "
                            "zichtbaar op deze straatfoto? Dit zijn ronde of rechthoekige putten "
                            "in de grond of bovengrondse containers op straat. "
                            "Antwoord alleen met JA of NEE."
                        )
                    }
                ]
            }]
        )
        antwoord = response.content[0].text.strip().upper()
        log.info(f"AI container-check: {antwoord}")
        return "JA" in antwoord
    except Exception as e:
        log.warning(f"AI container-check mislukt: {e}")
        return True  # Bij fout: gebruik huidige screenshot


def _tab_streetview(driver, wacht_klaar, nieuw_tabblad, coords, joined_coords, output_dir):
    from selenium.webdriver.common.by import By

    log.info("Tab 5: Google Street View")
    base_url = f"https://www.google.com/maps/@?api=1&map_action=pano&viewpoint={joined_coords}"

    # Navigeer en accepteer cookies
    nieuw_tabblad(base_url + "&heading=0", extra=2.0)
    for sel in [
        '//button[.//span[contains(text(),"Alles accepteren")]]',
        '//button[contains(text(),"Alles accepteren")]',
        '//form[contains(@action,"consent")]//button[last()]',
    ]:
        if _klik_als_aanwezig(driver, By.XPATH, sel, timeout=5):
            log.info("Street View: cookies geaccepteerd.")
            time.sleep(1.5)
            break

    # Refresh zodat Street View goed laadt (anders soms zwart scherm)
    driver.refresh()
    wacht_klaar()
    time.sleep(6)

    # Draai 360° rond en zoek container met AI (max 4 richtingen)
    richtingen = [0, 90, 180, 270]
    beste_screenshot = None

    for heading in richtingen:
        url = f"{base_url}&heading={heading}"
        driver.get(url)
        wacht_klaar()
        time.sleep(4)

        img = _screenshot_viewport(driver)
        tijdelijk_pad = os.path.join(output_dir, f"_streetview_tmp_{heading}.png")
        img.save(tijdelijk_pad)

        if _ai_ziet_container(tijdelijk_pad):
            log.info(f"Street View: container gevonden op heading {heading}°!")
            beste_screenshot = tijdelijk_pad
            break
        else:
            log.info(f"Street View: geen container op heading {heading}°, volgende richting...")
            beste_screenshot = tijdelijk_pad  # bewaar als fallback

    # Sla het beste screenshot op als definitieve Locatiefoto
    import shutil
    definitief = os.path.join(output_dir, "Locatiefoto.png")
    shutil.copy(beste_screenshot, definitief)

    # Verwijder tijdelijke bestanden
    for heading in richtingen:
        tmp = os.path.join(output_dir, f"_streetview_tmp_{heading}.png")
        if os.path.exists(tmp):
            os.remove(tmp)

    log.info("Street View: screenshot opgeslagen als Locatiefoto.png")


# ---------------------------------------------------------------------------
# 6. Afstandmeten.nl
# ---------------------------------------------------------------------------

def _tab_afstandmeten(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir):
    from selenium.webdriver.common.by import By
    from selenium.webdriver.common.keys import Keys
    from selenium.webdriver.common.action_chains import ActionChains

    log.info("Tab 6: Afstandmeten.nl")
    nieuw_tabblad("https://afstandmeten.nl/", extra=3.5)

    # Sluit alle cookiebanners en reclame
    cookie_selectors = [
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'toestemming')]"),
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'accepteer')]"),
        (By.XPATH, "//*[contains(translate(text(),'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'akkoord')]"),
        (By.CSS_SELECTOR, ".cc-btn, .cc-allow, button[id*='accept'], button[class*='consent']"),
        (By.XPATH, "//button[contains(@class,'close') or contains(@aria-label,'Sluit') or contains(@aria-label,'close')]"),
    ]
    for by, sel in cookie_selectors:
        _klik_als_aanwezig(driver, by, sel, timeout=3)

    # Adres invoeren linksboven
    try:
        veld = driver.find_element(By.ID, "qId")
        veld.clear()
        veld.send_keys(target_address)
        time.sleep(1.5)

        # Eerste suggestie klikken of Enter
        gevonden = _klik_als_aanwezig(
            driver, By.CSS_SELECTOR,
            ".pac-item:first-child, .ui-menu-item:first-child, "
            "[class*='suggestion']:first-child, [class*='result']:first-child",
            timeout=3
        )
        if not gevonden:
            veld.send_keys(Keys.RETURN)
    except Exception as e:
        log.warning(f"Afstandmeten: zoekveld: {e}")

    time.sleep(4)

    # Sluit eventuele reclame-overlays die na laden verschijnen
    for by, sel in cookie_selectors:
        _klik_als_aanwezig(driver, by, sel, timeout=2)

    # Vind de kaart en klik op het middelpunt (= de locatie)
    kaart = None
    for sel in ["#map", ".leaflet-container", ".map-container", "#mapCanvas", "[class*='map']"]:
        try:
            kaart = driver.find_element(By.CSS_SELECTOR, sel)
            if kaart:
                break
        except Exception:
            pass

    if kaart:
        size   = kaart.size
        midden_x = size["width"] // 2
        midden_y = size["height"] // 2

        # Klik op het middelpunt = start (container locatie)
        ActionChains(driver).move_to_element_with_offset(
            kaart, midden_x, midden_y
        ).click().perform()
        log.info("Afstandmeten: startpunt geplaatst (container locatie).")
        time.sleep(1.5)

        # Klik ~150px rechtsonder = eindpunt (~75m bij standaard zoom)
        ActionChains(driver).move_to_element_with_offset(
            kaart, midden_x + 150, midden_y + 80
        ).click().perform()
        log.info("Afstandmeten: eindpunt geplaatst (~75m).")
        time.sleep(2)
    else:
        log.warning("Afstandmeten: kaart niet gevonden, screenshot zonder meetlijn.")

    time.sleep(2)
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
    options.add_argument("--start-maximized")     # zichtbaar, zodat je kunt meekijken
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-popup-blocking")
    options.add_argument("--disable-notifications")
    options.add_argument("--lang=nl-NL")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)
    options.add_experimental_option("prefs", {
        "profile.default_content_setting_values.notifications": 2,
        "intl.accept_languages": "nl,nl-NL",
    })

    try:
        from webdriver_manager.chrome import ChromeDriverManager
        service = Service(ChromeDriverManager().install())
    except Exception:
        service = Service(executable_path=os.path.join(DEPS, "chromedriver.exe"))

    driver = webdriver.Chrome(service=service, options=options)

    def wacht_klaar(timeout=20):
        from selenium.webdriver.support.ui import WebDriverWait
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
        _tab_streetview(driver, wacht_klaar, nieuw_tabblad, coords, joined_coords, output_dir)
        _tab_afstandmeten(driver, wacht_klaar, nieuw_tabblad, target_address, output_dir)
    finally:
        driver.quit()
        log.info("Browser gesloten.")


# ---------------------------------------------------------------------------
# run_volledig
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

    log.info(f"=== Klaar! Excel staat op: {excel_pad} ===")
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
