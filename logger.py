"""
logger.py

Dieses Modul stellt eine Logging-Funktion bereit, um Ereignisse mit Zeitstempel, Dateinamen
und Zeilennummer zu protokollieren.

Die Logdatei wird bei jedem Programmstart neu erstellt.
Das Logging-Level kann über die `.env`-Datei gesteuert werden.
"""

import os
import datetime
import inspect
from config import LOG_FILE, DEBUG_LEVEL  # Importiert Logdateipfad und Debug-Level-Konfiguration

# Entfernt die alte Logdatei, falls sie existiert, um bei jedem Start eine neue zu erzeugen
if os.path.exists(LOG_FILE):
    os.remove(LOG_FILE)

def log(message, level=1):
    """
    Protokolliert eine Nachricht in die Logdatei mit Zeitstempel, Modulnamen und Zeilennummer.
    
    Parameter:
    - message (str): Die zu protokollierende Nachricht.
    - level (int): Die Wichtigkeitsstufe der Nachricht (je höher, desto detaillierter).
                   Nachrichten werden nur protokolliert, wenn `level` kleiner oder gleich dem
                   DEBUG_LEVEL aus der Konfiguration ist.
    
    Details:
    - Die Funktion ermittelt dynamisch Dateiname und Zeilennummer des Funktionsaufrufs, 
      um die Protokollierung zu präzisieren.
    - Nachrichten mit einer Wichtigkeitsstufe, die über dem aktuellen DEBUG_LEVEL liegen,
      werden ignoriert.
    """

    # Überspringt das Logging, wenn die Nachricht unter die aktuell konfigurierten Log-Ebenen fällt
    if level > DEBUG_LEVEL:
        return

    # Ermittelt die Aufrufinformationen (Dateiname und Zeilennummer), um den Ursprungsort zu identifizieren
    frame = inspect.currentframe().f_back
    filename = os.path.basename(frame.f_code.co_filename)  # Extrahiert nur den Dateinamen
    lineno = frame.f_lineno  # Zeilennummer des Log-Aufrufs

    # Erstellt einen Zeitstempel für die Logzeile
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Formatiert die Lognachricht mit Zeitstempel, Dateinamen und Zeilennummer
    log_entry = f"[{timestamp}] [{filename}:{lineno}] {message}"

    # Schreibt die Lognachricht in die Logdatei (Anhängen der neuen Einträge)
    with open(LOG_FILE, "a", encoding="utf-8") as log_file:
        log_file.write(log_entry + "\n")