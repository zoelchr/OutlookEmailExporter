"""
ui_loader.py

Dieses Modul ist zuständig für das dynamische Laden der grafischen Benutzeroberfläche
aus einer `.ui`-Datei (erstellt mit Qt Designer).

Die Klasse `MailGUI` stellt das zentrale GUI-Fenster dar, bindet alle relevanten Widgets
(z. B. Buttons, ComboBoxen, Tabellenansicht) und aktiviert die Statusleiste für Debug-Ausgaben.
"""

import os                                   # Ermöglicht den Zugriff auf das Dateisystem
from PySide6.QtUiTools import QUiLoader    # Lädt zur Laufzeit UI-Designs aus .ui-Dateien
from PySide6.QtWidgets import QMainWindow, QComboBox, QPushButton, QTableView  # GUI-Komponenten
from PySide6.QtCore import QFile, QObject, Qt  # Nützliche Klassen und Konstanten für GUI-Funktionen

from config import DEBUG_LEVEL             # Importiert die Konfiguration des Debug-Levels
from logger import log                     # Importiert die benutzerdefinierte Logging-Funktion

class MailGUI(QMainWindow):
    """
    Die Klasse `MailGUI` repräsentiert das Hauptfenster der Anwendung. 
    Sie lädt die Benutzeroberfläche zur Laufzeit und verwaltet zentrale GUI-Komponenten.
    """
    def __init__(self):
        super().__init__()

        # Initialisiert den UI-Loader
        loader = QUiLoader()  # Dynamisches Laden der .ui-Datei
        pfad_ui = os.path.join(os.path.dirname(__file__), "outlook_email_exporter.ui")
        log(f"📄 Lade UI-Datei: {pfad_ui}", level=2)

        # Öffnet und prüft die .ui-Datei auf Existenz
        ui_file = QFile(pfad_ui)
        if not ui_file.exists():
            log("❌ UI-Datei nicht gefunden!", level=0)  # Ausgabe bei fehlender Datei
            return

        # Lädt das UI aus der Datei
        ui_file.open(QFile.ReadOnly)
        self.ui = loader.load(ui_file, None)
        ui_file.close()

        # Prüft, ob das UI erfolgreich geladen wurde
        if not self.ui:
            log("❌ Fehler beim Laden der UI!", level=0)
            return

        # Setzt das geladene Layout als zentrales Widget des Fensters
        self.setCentralWidget(self.ui)
        log("✅ UI-Datei erfolgreich geladen", level=1)

        # Initialisiert die Statusleiste zur Anzeige von Debug-Informationen
        self.statusBar().showMessage(f"Debug-Level: {DEBUG_LEVEL}")
        log("📊 Statusleiste mit Debug-Level aktualisiert", level=3)

        # Helfermethode zum Binden von Widgets mit Protokollierung
        def bind_widget(widget_class, name, level=2):
            """
            Sucht ein Widget anhand seiner Klasse und seines Objekt-Namens und protokolliert das Ergebnis.
            """
            widget = self.ui.findChild(widget_class, name, Qt.FindChildrenRecursively)
            if widget:
                log(f"✅ Widget gebunden: {name}", level=level)
            else:
                log(f"⚠️ Widget NICHT gefunden: {name}", level=0)
            return widget

        # Bindet zentrale Widgets der Anwendung für spätere Interaktionen
        self.combo_postfach = bind_widget(QComboBox, "comboBox_Postfach")
        self.combo_verzeichnis = bind_widget(QComboBox, "comboBox_Verzeichnis")
        self.combo_exportziel = bind_widget(QComboBox, "comboBox_Exportziel")

        self.button_export_msg = bind_widget(QPushButton, "pushButton_Export_MSG")
        self.button_export_pdf = bind_widget(QPushButton, "pushButton_Export_PDF")
        self.button_exit = bind_widget(QPushButton, "pushButton_Exit")

        self.table_view = bind_widget(QTableView, "tableView_Emails")
        self.action_exit = bind_widget(QObject, "actionExit")

        # Ausgabe von Widgets für Debugging (optional, deaktiviert)
        # for child in self.ui.findChildren(QObject):
        #     print(child.objectName())
        
        log("🟢 MailGUI erfolgreich initialisiert", level=1)