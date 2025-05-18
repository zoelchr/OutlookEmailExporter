"""
outlook_email_exporter.py

Hauptskript zum Starten der Anwendung. Es initialisiert das Qt-Framework, lädt die GUI,
verbindet alle Steuersignale (über `gui_controller`) und startet die Ereignisschleife.

Dieses Modul dient als zentraler Einstiegspunkt für das gesamte Outlook-Export-Tool.
"""
from PySide6.QtWidgets import QApplication  # Importiert die QApplication-Klasse, die erforderlich ist, um eine Qt-Anwendung zu erstellen
from ui_loader import MailGUI               # Importiert das MailGUI-Hauptfenster aus dem Modul ui_loader
import sys                                  # Importiert das sys-Modul, um Zugriff auf Argumente und Systemfunktionen zu erhalten

from logger import log                      # Importiert die benutzerdefinierte Log-Funktion zur Protokollierung
from config import EXPORT_PATH              # Importiert den Exportpfad aus den Konfigurationsdateien

from gui_controller import connect_gui_signals

# Protokollierung des Programmstarts für Debugging und Nachvollziehbarkeit
log("Programmstart", level=1)
log("GUI wird geladen...", level=2)
log(f"Exportpfad laut .env: {EXPORT_PATH}", level=3)

# Hauptskriptblock: wird nur ausgeführt, wenn dieses Skript direkt gestartet wird
if __name__ == "__main__":
    # Erstellt eine Instanz der QApplication, die für die Durchführung jeder Qt-Anwendung erforderlich ist
    app = QApplication(sys.argv)
    
    # Initialisiert und zeigt das Hauptfenster der Anwendung an
    window = MailGUI()

    connect_gui_signals(window)

    window.show()
    
    # Startet die Ereignisschleife der GUI-Anwendung und beendet das Programm,
    # wenn alle Fenster geschlossen werden
    sys.exit(app.exec())