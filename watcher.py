# -*- coding: utf-8 -*-
"""
Hot-folder watcher — bewaakt de map `input_zone/` en verwerkt elk nieuw JSON-bestand
volledig automatisch: browser openen, screenshots maken, Excel genereren.

Gebruik:
    python watcher.py

Mappenstructuur (wordt automatisch aangemaakt):
    input_zone/   — drop hier je input.json bestanden
    verwerkt/     — succesvol verwerkte JSON's
    mislukt/      — JSON's die een fout veroorzaakten
    Pythonwerk/   — gegenereerde Excel-bestanden per gemeente/locatie
"""

import json
import logging
import os
import shutil
import sys
import time

from watchdog.events import FileSystemEventHandler
from watchdog.observers.polling import PollingObserver  # werkt altijd op Windows

FILE_LOCATION = os.path.dirname(os.path.realpath(__file__))
INPUT_DIR  = os.path.join(FILE_LOCATION, "input_zone")
DONE_DIR   = os.path.join(FILE_LOCATION, "verwerkt")
FAILED_DIR = os.path.join(FILE_LOCATION, "mislukt")

for d in (INPUT_DIR, DONE_DIR, FAILED_DIR):
    os.makedirs(d, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-8s  %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler(os.path.join(FILE_LOCATION, "watcher.log"), encoding="utf-8"),
    ],
)
log = logging.getLogger(__name__)

# Bijhouden welke bestanden al verwerkt zijn (voorkomt dubbele verwerking)
al_verwerkt = set()


def verwerk_json(json_pad: str):
    # Normaliseer pad zodat hoofdletters/slashes niet uitmaken op Windows
    json_pad = os.path.normpath(json_pad)
    bestandsnaam = os.path.basename(json_pad)

    if json_pad in al_verwerkt:
        return
    al_verwerkt.add(json_pad)

    log.info(f"Nieuw bestand gevonden: {bestandsnaam} — verwerking gestart")
    time.sleep(1.0)  # Wacht tot bestand volledig is weggeschreven

    if not os.path.exists(json_pad):
        log.warning(f"Bestand verdwenen voor verwerking: {bestandsnaam}")
        return

    try:
        with open(json_pad, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        log.error(f"JSON inlezen mislukt ({bestandsnaam}): {e}")
        _verplaats(json_pad, FAILED_DIR, bestandsnaam)
        return

    try:
        from auto_run import run_volledig
        excel_pad = run_volledig(data)
        log.info(f"Succes! Excel klaar: {excel_pad}")
        _verplaats(json_pad, DONE_DIR, bestandsnaam)
        log.info(f"{bestandsnaam} verplaatst naar verwerkt/")
    except Exception as e:
        import traceback
        log.error(f"Verwerking mislukt ({bestandsnaam}):\n{traceback.format_exc()}")
        _verplaats(json_pad, FAILED_DIR, bestandsnaam)
        log.warning(f"{bestandsnaam} verplaatst naar mislukt/")


def _verplaats(bron: str, doel_map: str, naam: str):
    if not os.path.exists(bron):
        return
    doel = os.path.join(doel_map, naam)
    if os.path.exists(doel):
        basis, ext = os.path.splitext(naam)
        doel = os.path.join(doel_map, f"{basis}_{int(time.time())}{ext}")
    shutil.move(bron, doel)


def scan_bestaande_bestanden():
    """Verwerk JSON-bestanden die al in input_zone lagen toen de watcher startte."""
    bestanden = [
        f for f in os.listdir(INPUT_DIR)
        if f.lower().endswith(".json")
    ]
    if bestanden:
        log.info(f"{len(bestanden)} bestaand(e) bestand(en) gevonden in input_zone — worden nu verwerkt.")
        for naam in bestanden:
            verwerk_json(os.path.join(INPUT_DIR, naam))
    else:
        log.info("Geen bestaande bestanden in input_zone. Wachten op nieuwe bestanden...")


class JsonHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".json"):
            verwerk_json(event.src_path)

    def on_modified(self, event):
        # Vangt ook op als een bestand wordt overschreven/gekopieerd
        if not event.is_directory and event.src_path.lower().endswith(".json"):
            verwerk_json(event.src_path)

    def on_moved(self, event):
        if not event.is_directory and event.dest_path.lower().endswith(".json"):
            verwerk_json(event.dest_path)


if __name__ == "__main__":
    log.info("=" * 55)
    log.info("  Productblad Watcher — B-Advice")
    log.info("=" * 55)
    log.info(f"Bewaakt map: {INPUT_DIR}")
    log.info("Drop een input.json in input_zone/ om te starten.")
    log.info("Stop met Ctrl+C.")
    log.info("-" * 55)

    # Verwerk bestanden die er al lagen
    scan_bestaande_bestanden()

    # Start de watcher (PollingObserver = meest betrouwbaar op Windows)
    observer = PollingObserver(timeout=2)  # controleert elke 2 seconden
    observer.schedule(JsonHandler(), INPUT_DIR, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Watcher gestopt door gebruiker.")
        observer.stop()
    observer.join()
