"""
outlook_email_exporter.py

Dieses Skript ist der Einstiegspunkt für die Outlook Email Exporter-Anwendung.
Es initialisiert die grafische Benutzeroberfläche (GUI) mithilfe von PySide6 und startet das Hauptfenster,
mit dem Benutzer E-Mails aus Outlook exportieren können.
"""

from PySide6.QtWidgets import QApplication  # Importiert die QApplication-Klasse, die erforderlich ist, um eine Qt-Anwendung zu erstellen
from ui_loader import MailGUI               # Importiert das MailGUI-Hauptfenster aus dem Modul ui_loader
import sys                                  # Importiert das sys-Modul, um Zugriff auf Argumente und Systemfunktionen zu erhalten

from logger import log                      # Importiert die benutzerdefinierte Log-Funktion zur Protokollierung
from config import EXPORT_PATH              # Importiert den Exportpfad aus den Konfigurationsdateien

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
    window.show()
    
    # Startet die Ereignisschleife der GUI-Anwendung und beendet das Programm,
    # wenn alle Fenster geschlossen werden
    sys.exit(app.exec())