import logging
logger = logging.getLogger(__name__)
"""
ui_loader.py


"""

import os  # Ermöglicht den Zugriff auf das Dateisystem
from PySide6.QtUiTools import QUiLoader  # Lädt zur Laufzeit UI-Designs aus .ui-Dateien
from PySide6.QtWidgets import QMainWindow, QComboBox, QPushButton, QTableView  # GUI-Komponenten
from PySide6.QtCore import QFile, QObject, Qt  # Nützliche Klassen und Konstanten für GUI-Funktionen
from outlook_email_exporter_ui import Ui_MainWindow  # Importieren des generierten Codes

from config import DEBUG_LEVEL  # Importiert die Konfiguration des Debug-Levels

LEVEL_NAME_MAP = {
    0: "ERROR",
    1: "WARNING",
    2: "INFO",
    3: "DEBUG",
}

class MyMainWindow(QMainWindow):
    """
    Die Klasse `MailGUI` repräsentiert das Hauptfenster der Anwendung.
    Sie lädt die Benutzeroberfläche zur Laufzeit und verwaltet zentrale GUI-Komponenten.
    """

    def __init__(self):
        super().__init__()
        logger.info(f"MyMainWindow init gestartet.")

        # Die auf Basis Qt Designer generierte UI-Klasse
        self.ui = Ui_MainWindow()  # Instanz der generierten UI-Klasse erstellen
        self.ui.setupUi(self)  # Die GUI auf das aktuelle Fenster anwenden
        logger.info(f"Die generierte GUI wurde auf das aktuelle MyMainWindow angewendet.")

        # Initialisiert die Statusleiste zur Anzeige von Debug-Informationen
        level_name = LEVEL_NAME_MAP.get(DEBUG_LEVEL, f"Unbekannt ({DEBUG_LEVEL})")
        self.statusBar().showMessage(f"Debug-Level: {level_name}.")
        logger.debug("Statusleiste mit Debug-Level aktualisiert.")

        # Helfermethode zum Binden von Widgets mit Protokollierung
        def bind_widget(widget_class, name, level=2):
            """
            Sucht ein Widget anhand seiner Klasse und seines Objekt-Namens und protokolliert das Ergebnis.
            """
            widget = self.ui.findChild(widget_class, name, Qt.FindChildrenRecursively)
            if widget:
                logger.debug(f"✅ Widget gebunden: {name}")
            else:
                logger.warning(f"⚠️ Widget NICHT gefunden: {name}")
            return widget

        # # Bindet zentrale Widgets der Anwendung für spätere Interaktionen
        # self.combo_postfach = bind_widget(QComboBox, "comboBox_Postfach")
        # self.combo_verzeichnis = bind_widget(QComboBox, "comboBox_Verzeichnis")
        # self.combo_exportziel = bind_widget(QComboBox, "comboBox_Exportziel")
        #
        # self.button_export_msg = bind_widget(QPushButton, "pushButton_Export_MSG")
        # self.button_export_pdf = bind_widget(QPushButton, "pushButton_Export_PDF")
        # self.button_exit = bind_widget(QPushButton, "pushButton_Exit")
        #
        # self.table_view = bind_widget(QTableView, "tableView_Emails")
        # self.action_exit = bind_widget(QObject, "actionExit")
        #
        # # Ausgabe von Widgets für Debugging
        # logger.debug("Gefundene QObjects:")
        # for child in self.ui.findChildren(QObject):
        #     cls_name = child.metaObject().className()
        #     obj_name = child.objectName()
        #     logger.debug(f"- {cls_name}: {obj_name if obj_name else '(kein Name)'}")

        logger.info("🟢 MailGUI erfolgreich initialisiert")