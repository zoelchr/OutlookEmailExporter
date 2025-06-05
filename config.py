"""
config.py

Dieses Modul liest Konfigurationswerte aus einer `.env`-Datei ein und stellt sie zentral
für andere Anwendungsteile zur Verfügung.

Unterstützt werden:
- einstellbare Debug-Stufen,
- dynamische Definition des Log-Dateinamens,
- Exportverzeichnis für Ausgabedateien,
- eine Liste von Postfächern, die in der GUI nicht angezeigt werden sollen,
- eine Liste von Ordnernamen, die beim Laden der Outlook-Ordnerstruktur ignoriert werden sollen.
"""
import logging
app_logger = logging.getLogger(__name__)

import os
from dotenv import load_dotenv

# .env-Datei laden (sofern vorhanden)
load_dotenv()

# Allgemeine Einstellungen
DEBUG_LEVEL = int(os.getenv("DEBUG_LEVEL", "0"))
LOG_FILE = os.getenv("LOG_FILE", "debug.log")
EXPORT_PATH = os.getenv("EXPORT_PATH", "./export")

# Liste von Postfächern, die in der GUI nicht angezeigt werden sollen
IGNORE_POSTFAECHER = [s.strip() for s in os.getenv("IGNORE_POSTFAECHER", "").split(",") if s.strip()]

# Liste von Begriffen, bei denen ein Ordnername ausgeschlossen wird (case-insensitive)
# Beispiel: "Kalender, Kontakte" → schließt alle Ordner aus, die diese Begriffe im Namen tragen
EXCLUDE_FOLDERNAMES = [s.strip().lower() for s in os.getenv("EXCLUDE_FOLDERNAMES", "").split(",") if s.strip()]

# Maximale Suchtiefe nach Ordnern (0 = nur Postfach-Root, 1 = Postfach-Root + 1 Ebene, ...)
MAX_FOLDER_DEPTH = int(os.getenv("MAX_FOLDER_DEPTH", "1"))

# Maximale Anzahl Einträge in der Exportziel-Combobox
MAX_SAVED_EXPORT_TARGETS = int(os.getenv("MAX_SAVED_EXPORT_TARGETS", "20"))
MAX_ENV_EXPORT_TARGETS = int(os.getenv("MAX_ENV_EXPORT_TARGETS", "5"))

# Liste mit den bekannten Absendern
KNOWNSENDER_FILE = os.getenv("KNOWNSENDER_FILE", "./known_senders_private.csv")

# Testmodus (kein Email-Export)
TEST_MODE = os.getenv("TEST_MODE", "false").lower() in ["true", "1", "yes", "y"]