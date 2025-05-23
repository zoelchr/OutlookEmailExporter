"""
ui_loader.py


"""
import os  # Ermöglicht den Zugriff auf das Dateisystem

from PySide6.QtUiTools import QUiLoader  # Lädt zur Laufzeit UI-Designs aus .ui-Dateien
from PySide6.QtWidgets import QMainWindow, QComboBox, QPushButton, QTableView  # GUI-Komponenten
from PySide6.QtCore import QFile, QObject, Qt  # Nützliche Klassen und Konstanten für GUI-Funktionen
from ui_main_window import Ui_MainWindow  # Importieren des generierten Codes

import logging
logger = logging.getLogger(__name__)

from logger import LEVEL_NAME_MAP  # LEVEL_NAME_MAP korrekt importieren
from config import DEBUG_LEVEL, MAX_FOLDER_DEPTH  # Importiere relevante Konfigurationen


class MyMainWindow(QMainWindow):
    """
    Die Klasse `MyMainWindow` repräsentiert das Hauptfenster der Anwendung.


    """
    def __init__(self):
        # Helfermethode zum Binden von Widgets mit Protokollierung
        def bind_widget(widget_class, name, level=2):
            """
            Sucht ein Widget anhand seiner Klasse und seines Objekt-Namens und protokolliert das Ergebnis.
            """
            widget = self.findChild(widget_class, name, Qt.FindChildrenRecursively)
            if widget:
                logger.debug(f"Widget gebunden: {name}")
            else:
                logger.warning(f"Widget NICHT gefunden: {name}")
            return widget

        super().__init__()
        logger.debug(f"MyMainWindow init gestartet.")

        # Die auf Basis Qt Designer generierte UI-Klasse
        self.ui = Ui_MainWindow()  # Instanz der generierten UI-Klasse erstellen
        self.ui.setupUi(self)  # Die GUI auf das aktuelle Fenster anwenden
        logger.debug(f"Die generierte GUI wurde auf das aktuelle MyMainWindow angewendet.")

        # Initialisiert die Titelleiste des Hauptfensters
        self.setWindowTitle("Outlook Email Exporter")
        logger.debug("Titelleiste des Hauptfensters initialisiert.")

        # Initialisiert die Statusleiste zur Anzeige von Debug-Informationen
        level_name = LEVEL_NAME_MAP.get(DEBUG_LEVEL, f"Unbekannt ({DEBUG_LEVEL})")
        self.statusBar().showMessage(f"Debug-Level: {level_name} | Max. Suchtiefe für Ordnerabfrage: {MAX_FOLDER_DEPTH}")
        logger.debug("Statusleiste mit Debug-Level aktualisiert.")

        # Bindet zentrale Widgets der Anwendung für spätere Interaktionen
        self.combo_postfach = bind_widget(QComboBox, "comboBox_Postfach")
        self.combo_verzeichnis = bind_widget(QComboBox, "comboBox_Verzeichnis")
        self.combo_exportziel = bind_widget(QComboBox, "comboBox_Exportziel")

        self.button_export_msg = bind_widget(QPushButton, "pushButton_Export_MSG")
        self.button_export_pdf = bind_widget(QPushButton, "pushButton_Export_PDF")
        self.button_exit = bind_widget(QPushButton, "pushButton_Exit")

        self.table_view = bind_widget(QTableView, "tableView_Emails")
        self.action_exit = bind_widget(QObject, "actionExit")

        # Ausgabe von Widgets für Debugging
        logger.debug("Gefundene QObjects:")
        for child in self.findChildren(QObject):
            cls_name = child.metaObject().className()
            obj_name = child.objectName()
            logger.debug(f"- {cls_name}: {obj_name if obj_name else '(kein Name)'}")

        logger.debug("MyMainWindow erfolgreich initialisiert.")

