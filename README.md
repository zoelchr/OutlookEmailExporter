# Outlook Email Exporter

## Übersicht

Dieses Projekt bietet ein Tool zur Verarbeitung und zum Export von E-Mails, das mit PySide6 für die Benutzeroberfläche und weiteren Modulen wie pywin32 und jinja2 arbeitet. Es wurde entwickelt, um E-Mails effizient zu analysieren, Metadaten zu extrahieren und diese in verschiedenen Formaten zu exportieren.

Die generierte ausführbare Datei ermöglicht die Verteilung des Tools ohne Python-Installation auf den Zielsystemen. Hierbei wird PyInstaller verwendet, um den Build-Prozess durchzuführen.

---

## Features

- GUI-Erstellung mit PySide6: Ein benutzerfreundliches Interface, um E-Mails zu verwalten und zu exportieren.
- Integration mit Microsoft Outlook: Für den Zugriff auf E-Mails (über pywin32).
- Robuste Verarbeitung von MSG-Dateien: Metadatenextraktion, Formatierung und Export.
- File-Handling: Anpassung und Sicherstellung der Kompatibilität bei Exporten.
- Log-Management: Protokollierung aller Aktivitäten für Debugging und Nachvollziehbarkeit.

---

## Voraussetzungen

Bevor du das Projekt baust oder ausführst, stelle sicher, dass die folgenden Voraussetzungen erfüllt sind:

### Software:

- Python-Version: 3.10 oder höher
- System: Windows 10 oder höher
- VirtualEnv: Wird für die Isolierung der Abhängigkeiten genutzt.

### Abhängigkeiten:

Installiere die benötigten Python-Pakete, indem du den folgenden Befehl ausführst:
    
    pip install -r requirements.txt

### Hauptpakete:

- PySide6: Für die grafische Benutzeroberfläche.
- pywin32: Zugriff auf Outlook- und Windows-Dienste.
- jinja2: Template-Handling.
- pandas: Verarbeitung und Export von tabellarischen Daten.

---

## Build-Prozess mit PyInstaller

Die Anwendung wird mit PyInstaller in eine ausführbare .exe verpackt, um die Verteilung zu erleichtern. Der Build integriert notwendige Python-Module und Abhängigkeiten direkt in die ausführbare Datei.

### Befehl zum Erstellen der ausführbaren Datei:

    pyinstaller --onefile --console main.py --name outlookemailexporter --hidden-import win32timezone

### Parameter-Beschreibung:

- --onefile: Alle Dateien werden in eine einzige .exe gepackt.
- --console: Die Konsole wird beim Ausführen der .exe angezeigt (gut für Debugging).
- --name outlookemailexporter: Gibt der generierten .exe den Namen outlookemailexporter.
- --hidden-import win32timezone: Behebt fehlende Module wie win32timezone, die PyInstaller nicht automatisch erkennt.

---

## Anwendung ausführen

Nach der Erstellung:

Die generierte Datei befindet sich im Ordner dist und kann wie jede andere ausführbare Datei gestartet werden:

    dist/outlookemailexporter.exe

Beim Start öffnet sich die grafische Benutzeroberfläche.

---

## Häufige Probleme und Lösungen

1. Fehlermeldung: "No module named 'win32timezone'"
   - Ursache: Das Modul wurde von PyInstaller während des Builds nicht erkannt.
   - Lösung: Stelle sicher, dass --hidden-import win32timezone im Build-Befehl enthalten ist.

2. Fehlender Zugriff auf Outlook
   - Ursache: pywin32 ist nicht korrekt installiert.
   - Lösung: Führe folgenden Befehl aus:
   
        pip install pywin32
        python Scripts/pywin32_postinstall.py -install

3. jinja2 nicht gefunden:
   - Ursache: Abhängigkeit fehlt.
   - Lösung: Installiere es über:
   
        pip install jinja2

---

## Verzeichnisse und Dateien

### Wichtige Dateien im Projekt:

- main.py: Der Haupteinstiegspunkt der Anwendung.
- export_manager.py: Beinhaltet die Logik für das Exportieren der E-Mails.
- config.py: Konfigurationsdatei für den Exportpfad und andere Einstellungen.
- .env: Definiert Umgebungsvariablen, z. B. DEBUG_LEVEL und EXPORT_PATH.

### Ordner:

- dist/: Enthält die generierte .exe nach dem Build.
- build/: Temporäre Dateien während des Build-Prozesses (kann gelöscht werden).
- .venv/: Virtuelle Umgebung mit isolierten Abhängigkeiten (falls genutzt).

---

## Konfiguration

Die Datei .env erlaubt es, verschiedene Einstellungen für die Anwendung zu definieren:

Beispiel:

    # Logging-Stufe: 0=ERROR, 1=WARNING, 2=INFO, 3=DEBUG, 4=TRACE
    DEBUG_LEVEL=3
    # Logging-Ausgabe zusätzlich auf Console
    LOG_TO_CONSOLE=false
    # Pfad zur Log-Datei
    LOG_FILE=debug.log
    # Exportpfad
    EXPORT_PATH=./logs

---

## Unterstützung und Fragen

Falls während der Nutzung oder beim Build-Prozess Probleme auftreten, melde dich gerne! 😊