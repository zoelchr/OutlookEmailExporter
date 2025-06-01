"""
logger.py

Dieses Modul konfiguriert die Logging-Funktionalität der Anwendung basierend auf den Einstellungen in der `.env`-Datei.
Es stellt sicher, dass Protokollnachrichten in einer Datei sowie optional in der Konsole angezeigt werden.
Zudem wird eine benutzerdefinierte TRACE-Ebene hinzugefügt, die besonders detaillierte Debugging-Informationen bereitstellt.

### Hauptaufgaben:
1. **Logging-Konfiguration**:
   - Legt das Log-Level basierend auf der Umgebungsvariablen `DEBUG_LEVEL` fest.
   - Unterstützte Stufen: ERROR, WARNING, INFO, DEBUG, TRACE.

2. **Dateibasiertes Logging**:
   - Alle Log-Einträge werden standardmäßig in der `.env`-Datei spezifizierten Datei (`LOG_FILE`) gespeichert.
   - Vor jedem Start wird die bestehende Log-Datei gelöscht.

3. **Optionales Konsolen-Logging**:
   - Kann über die Umgebungsvariable `LOG_TO_CONSOLE` aktiviert werden.
   - Gibt die Log-Einträge zusätzlich in der Konsole aus.

4. **TRACE-Level**:
   - Eine neue Logging-Stufe für detaillierte Debugging-Meldungen, dessen Schweregrad unterhalb von DEBUG liegt.

5. **Erweiterbarkeit**:
   - Benutzerdefinierte Logging-Funktion (`trace`) an die `Logger`-Klasse hinzugefügt.

### Übersicht: Funktionen & Konfiguration
1. **TRACE-Level hinzufügen**
   Definiert ein neues Logging-Level mit der Bezeichnung "TRACE" und dem Wert 9 (unterhalb von DEBUG).

2. **Handler entfernen**
   Entfernt alle existierenden Logging-Handler, um sicherzustellen, dass die neue Konfiguration gültig ist.

3. **Log-Datei löschen**
   Löscht die bestehende Log-Datei beim Start der Anwendung, um eine saubere Protokollierung sicherzustellen.

4. **`basicConfig`**
   Richtet die Grundkonfiguration für das Logging ein, einschließlich Format, Level und Ausgabedatei.

5. **TRACE-Methode definieren**
   Fügt die Methode `trace` als benutzerdefinierte Methode zur `Logger`-Klasse hinzu.

6. **Konsolen-Handler hinzufügen**
   Aktiviert das optionale Konsolen-Logging, wenn `LOG_TO_CONSOLE` aktiv ist.

### Einschränkungen:
- Alte Log-Dateien werden automatisch gelöscht.
- Das Modul benötigt die Werte `LOG_FILE` und `DEBUG_LEVEL` aus der `.env`-Datei.

### Beispiele:
- Aktivierung von TRACE-Logs:
  Setzen Sie die Umgebungsvariable `DEBUG_LEVEL` auf 4.
- Konsolen-Logging aktivieren:
  Setzen Sie die Umgebungsvariable `LOG_TO_CONSOLE` auf `true`.

"""
import os
import logging
from config import LOG_FILE, DEBUG_LEVEL

# Konsolen-Logging zusätzlich aktivieren
log_to_console = os.getenv("LOG_TO_CONSOLE", "false").lower() == "true"

# Vorherige Handler vollständig entfernen, damit basicConfig garantiert greift
for handler in logging.root.handlers[:]:
    logging.root.removeHandler(handler)

# Füge eine neue TRACE Logging-Stufe hinzu
TRACE_LEVEL = 9  # Ein Wert kleiner als DEBUG (10)
logging.addLevelName(TRACE_LEVEL, "TRACE")

# Mappe benutzerdefinierten DEBUG_LEVEL auf echtes logging-Level
LEVEL_NAME_MAP = {
    0: logging.ERROR,
    1: logging.WARNING,
    2: logging.INFO,
    3: logging.DEBUG,
    4: TRACE_LEVEL
}
log_level = LEVEL_NAME_MAP.get(DEBUG_LEVEL, logging.INFO)

# 🗑 Alte Log-Datei löschen (bei jedem Start)
try:
    if os.path.exists(LOG_FILE):
        os.remove(LOG_FILE)
except Exception as e:
    print(f"Konnte alte Logdatei nicht löschen: {e}")

# Logging-Grundkonfiguration (Datei)
logging.basicConfig(
    level=log_level,
    format='[%(asctime)s] [%(levelname)s] [%(name)s:%(lineno)d] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S',
    filename=LOG_FILE,
    filemode='a',  # Datei neu erzeugen (weil oben gelöscht)
    encoding='utf-8'
)

# TRACE-Methode für den Logger definieren
def trace(self, message, *args, **kwargs):
    if self.isEnabledFor(TRACE_LEVEL):
        self._log(TRACE_LEVEL, message, args, **kwargs)

# TRACE-Level zur Logger-Klasse hinzufügen
logging.Logger.trace = trace


# Konsolen-Logging zusätzlich aktivieren
if log_to_console:
    console_handler = logging.StreamHandler()
    console_handler.setLevel(log_level)
    console_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    logging.getLogger().addHandler(console_handler)

# Testausgabe – wird in Logdatei UND Konsole erscheinen
logging.getLogger(__name__).debug("Logging-Konfiguration abgeschlossen")
