# logger.py

import os
import logging
from config import LOG_FILE, DEBUG_LEVEL

# Konsolen-Logging zusätzlich aktivieren
log_to_console = os.getenv("LOG_TO_CONSOLE", "false").lower() == "true"

# 🔁 Vorherige Handler vollständig entfernen, damit basicConfig garantiert greift
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

# 🔢 Mappe benutzerdefinierten DEBUG_LEVEL auf echtes logging-Level
LEVEL_MAP = {
    0: logging.ERROR,
    1: logging.WARNING,
    2: logging.INFO,
    3: logging.DEBUG,
}
log_level = LEVEL_MAP.get(DEBUG_LEVEL, logging.INFO)

# 🗑 Alte Log-Datei löschen (bei jedem Start)
try:
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
except Exception as e:
    print(f"⚠️ Konnte alte Logdatei nicht löschen: {e}")

# 📝 Logging-Grundkonfiguration (Datei)
logging.basicConfig(
    level=log_level,
    format='[%(asctime)s] [%(levelname)s] [%(name)s:%(lineno)d] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    filename=LOG_FILE,
    filemode='a',  # Datei neu erzeugen (weil oben gelöscht)
    encoding='utf-8'
)

# 🖥 Konsolen-Logging zusätzlich aktivieren
if log_to_console:
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    logging.getLogger().addHandler(console_handler)

# 🧪 Testausgabe – wird in Logdatei UND Konsole erscheinen
logging.getLogger(__name__).debug("🧪 Logging-Konfiguration abgeschlossen")
