# -*- coding: utf-8 -*-
"""
Hot-folder watcher — bewaakt de map `input_zone/` en verwerkt elk nieuw JSON-bestand
automatisch via productblad_core.generate_excel().

Gebruik:
    pip install watchdog
    python watcher.py

Mappenstructuur die automatisch wordt aangemaakt:
    input_zone/   — drop hier je input.json bestanden
    verwerkt/     — verwerkte JSON's worden hierheen verplaatst
    mislukt/      — JSON's die een fout veroorzaakten
"""

import json
import logging
import os
import shutil
import sys
import time

from watchdog.events import FileSystemEventHandler
from watchdog.observers import Observer

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


def verwerk_json(json_pad: str):
    bestandsnaam = os.path.basename(json_pad)
    log.info(f"Nieuw bestand gedetecteerd: {bestandsnaam}")

    # Wacht kort zodat het bestand volledig is geschreven
    time.sleep(0.5)

    try:
        with open(json_pad, "r", encoding="utf-8") as f:
            data = json.load(f)
    except Exception as e:
        log.error(f"JSON inlezen mislukt ({bestandsnaam}): {e}")
        _verplaats(json_pad, FAILED_DIR, bestandsnaam)
        return

    try:
        from productblad_core import generate_excel
        uitvoerpad = generate_excel(data)
        log.info(f"Excel aangemaakt: {uitvoerpad}")
        _verplaats(json_pad, DONE_DIR, bestandsnaam)
        log.info(f"{bestandsnaam} verplaatst naar verwerkt/")
    except Exception as e:
        import traceback
        log.error(f"Generatie mislukt ({bestandsnaam}):\n{traceback.format_exc()}")
        _verplaats(json_pad, FAILED_DIR, bestandsnaam)


def _verplaats(bron: str, doel_map: str, naam: str):
    doel = os.path.join(doel_map, naam)
    # Voorkom overschrijven bij dubbele namen
    if os.path.exists(doel):
        basis, ext = os.path.splitext(naam)
        doel = os.path.join(doel_map, f"{basis}_{int(time.time())}{ext}")
    shutil.move(bron, doel)


class JsonHandler(FileSystemEventHandler):
    def on_created(self, event):
        if not event.is_directory and event.src_path.lower().endswith(".json"):
            verwerk_json(event.src_path)

    def on_moved(self, event):
        # Vangt ook bestanden op die in de map worden gesleept via verkenner
        if not event.is_directory and event.dest_path.lower().endswith(".json"):
            verwerk_json(event.dest_path)


if __name__ == "__main__":
    log.info(f"Watcher gestart. Bewaakt: {INPUT_DIR}")
    log.info("Stop met Ctrl+C.")

    observer = Observer()
    observer.schedule(JsonHandler(), INPUT_DIR, recursive=False)
    observer.start()

    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        log.info("Watcher gestopt.")
        observer.stop()
    observer.join()
