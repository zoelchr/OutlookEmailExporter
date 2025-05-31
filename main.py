"""
outlook_email_exporter.py

Dieses Modul ist der zentrale Einstiegspunkt für das Outlook-Export-Tool. Es initialisiert
die Hauptanwendung, richtet das Qt-Framework ein, lädt die grafische Benutzeroberfläche (GUI)
und verbindet die Steuerlogik (z. B. Signal-Slot-Mechanismen) durch den `gui_controller`.

Funktionen und Aufgaben:
-------------------------
- Initialisierung der Qt-Anwendung und Ladeprozess der GUI-Komponenten.
- Verbindung der Steuerlogik mit der Benutzeroberfläche (z. B. Listen oder Auswahlfelder).
- Starten der Haupt-Ereignisschleife, um Interaktionen und Ereignisse in Echtzeit zu verarbeiten.

Besonderheiten:
---------------
- Das Modul übernimmt die zentrale Aufgabe, verschiedene Teile der Anwendung (z. B. Konfiguration,
  Export-Manager, Zielverzeichnisse) zu initialisieren und miteinander zu verbinden.
- Enthält keine direkte Geschäftslogik, sondern dient ausschließlich der Anwendungssteuerung und dem Startprozess.
- Unterstützt die Integration mit anderen Modulen wie `config.py`, `export_manager.py`, und `exportziel_manager.py`.

Hinweise:
---------
Weitere Anpassungen und Konfigurationen der App (z. B. Pfadangaben, Debug-Stufen) können über das
`config.py`-Modul vorgenommen werden. Alle relevanten Signale und Logiken werden durch die Instanz
des `gui_controller` koordiniert.
"""
from PySide6.QtWidgets import QApplication      # Importiert die QApplication-Klasse, die erforderlich ist, um eine Qt-Anwendung zu erstellen
import sys                                      # Importiert das sys-Modul, um Zugriff auf Argumente und Systemfunktionen zu erhalten

from config import EXPORT_PATH                  # Importiert den Exportpfad aus den Konfigurationsdateien

from ui_main_window import Ui_MainWindow      # Importieren das Ui_MainWindow-Hauptfenster aus dem Modul des generierten Codes
from my_main_window import MyMainWindow         # Importiert das MyMainWindow-Hauptfenster aus dem Modul ui_loader
from gui_controller import connect_gui_signals
from exportziel_manager import get_exportziel_manager

import logging
app_logger = logging.getLogger(__name__)  # dein echter Logger

# Protokollierung des Programmstarts für Debugging und Nachvollziehbarkeit
app_logger.debug(f"Exportpfad laut .env: {EXPORT_PATH}")
app_logger.debug("Programmstart...")

# Hauptskriptblock: wird nur ausgeführt, wenn dieses Skript direkt gestartet wird
if __name__ == "__main__":

    # Hier wird eine Instanz der QApplication erzeugt.
    # - ist das zentrale Objekt einer PySide-Anwendung. Es steuert das grundlegende Event-Handling und die Lebensdauer der GUI-Applikation. `QApplication`
    # - Es wird das Event-System von Qt initialisiert und auf die Kommandozeilenargumente `sys.argv` übergeben.
    # - Statt sys.argv kann auch eine leere Liste [] übergeben werden
    app = QApplication(sys.argv)
    app_logger.debug(f"Die Instanz QApplication wurde erzeugt.")

    # - Es wird eine Instanz von MyMainWindow erzeugt. Dies ist das Hauptfenster der Anwendung.
    # - MyMainWindow basiert auf der Qt-Klasse und enthält alle wichtigen Elemente der Benutzeroberfläche (z. B. Menüs, Widgets etc.).
    window = MyMainWindow()
    app_logger.debug(f"MyMainWindow wurde instanziiert.")

    # Verknüpfung der GUI-Komponenten mit Logik
    # - Die Funktion durchläuft alle relevanten Widgets der Klasse MyMainWindow (repräsentiert durch `window`).
    # - Für jedes relevante Signal der GUI-Komponenten verbindet sie ein Signal (`.clicked`, etc.) mit einer passenden Funktion (Slot).
    connect_gui_signals(window)

    # Initialisierung des ExportzielManagers (Singleton mit GUI binden)
    exportziel_manager = get_exportziel_manager(window)
    # Optional: Prüfen, ob die Instanz initialisiert wurde
    # Exportziele in der ComboBox initialisieren
    exportziel_manager.initExportTargets()  # Nach Instanziierung aufrufen

    app_logger.debug(f"ExportzielManager erfolgreich initialisiert: {exportziel_manager is not None}")


    # - Mit show() wird das Hauptfenster sichtbar gemacht und auf dem Bildschirm angezeigt.
    # - Es handelt sich um die Aufforderung an Qt, das Hauptfenster in das GUI-Ereignissystem zu integrieren und es anzuzeigen.
    window.show()
    app_logger.debug(f"MyMainWindow wird angezeigt.")

    # - Die Methode app.exec() startet die Haupt-Ereignisschleife (Main Event Loop) der Qt-Anwendung. Diese Schleife ist verantwortlich für:
    #     - Das Weiterleiten von Benutzerinteraktionen (wie Mausklicks, Tastatureingaben) an die Anwendung.
    #     - Das zeitgesteuerte Aktualisieren und Neuzeichnen der GUI.
    #
    # - Sie blockiert den Programmfluss, bis die Anwendung beendet oder geschlossen wird.
    # - Die Rückgabe von app.exec() beschreibt den Exit-Code der Anwendung, und diesen übergibt das Programm an sys.exit().
    sys.exit(app.exec())