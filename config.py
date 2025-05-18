"""
config.py

Dieses Modul liest Konfigurationswerte aus einer `.env`-Datei ein und stellt sie zentral
für andere Anwendungsteile zur Verfügung. Es unterstützt unter anderem:
- einstellbare Debug-Stufen,
- dynamische Definition des Log-Dateinamens,
- Exportverzeichnis für Ausgabedateien,
- eine Liste von Postfächern, die bei der Auswahl ignoriert werden sollen.
"""

import os
from dotenv import load_dotenv

# Lädt Umgebungsvariablen aus einer .env-Datei
load_dotenv()

# Konfigurationswerte mit Fallback auf Standardwerte
DEBUG_LEVEL = int(os.getenv("DEBUG_LEVEL", "0"))
LOG_FILE = os.getenv("LOG_FILE", "debug.log")
EXPORT_PATH = os.getenv("EXPORT_PATH", "./export")

# Kommagetrennte Liste von Postfächern, die nicht angezeigt werden sollen
IGNORE_POSTFAECHER = [s.strip() for s in os.getenv("IGNORE_POSTFAECHER", "").split(",") if s.strip()]
