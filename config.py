"""
config.py

Dieses Modul liest Konfigurationsdaten aus einer `.env`-Datei ein und stellt sie der Anwendung zur Verfügung.
"""

import os                                   # Ermöglicht den Zugriff auf Umgebungsvariablen
from dotenv import load_dotenv             # Lädt Umgebungsvariablen aus einer .env-Datei

# Lädt die Umgebungsvariablen aus der .env-Datei
load_dotenv()

# Konfigurationsparameter mit Standardwerten, falls keine Werte in der .env-Datei gefunden wurden
DEBUG_LEVEL = int(os.getenv("DEBUG_LEVEL", "0"))  # Debug-Level zur Steuerung der Protokollausgabe
LOG_FILE = os.getenv("LOG_FILE", "debug.log")    # Name der Logdatei
EXPORT_PATH = os.getenv("EXPORT_PATH", "./export")  # Standard-Exportpfad für gespeicherte Dateien
IGNORE_POSTFAECHER = [s.strip() for s in os.getenv("IGNORE_POSTFAECHER", "").split(",") if s.strip()]