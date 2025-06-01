"""
Die Datei `my_main_window.py` enthält die Hauptlogik der Anwendung und definiert das
Hauptfenster der GUI als Klasse `MyMainWindow`. Das Fenster basiert auf PySide6 und
bindet dynamisch Widgets, die über Qt Designer erstellt wurden. 

Hauptaufgaben dieser Datei:
- Initialisierung und Anwendung der generierten Benutzeroberfläche (UI).
- Bindung und Anpassung zentraler Widgets (z. B. Buttons, Komboboxen, Tabellenansicht).
- Anzeige von Debug- und Statusinformationen in der Statusleiste.
- Protokollierung und Prüfung der UI-Komponenten zur Unterstützung von Debugging und Fehlersuche.

Diese Struktur ermöglicht eine flexible und modulare Erweiterung der GUI-Funktionalität.
"""
import os  # Ermöglicht den Zugriff auf das Dateisystem
from PySide6.QtUiTools import QUiLoader  # Lädt zur Laufzeit UI-Designs aus .ui-Dateien
from PySide6.QtWidgets import QMainWindow, QComboBox, QPushButton, QTableView, QDialog, QCheckBox  # GUI-Komponenten
from PySide6.QtCore import QFile, QObject, Qt  # Nützliche Klassen und Konstanten für GUI-Funktionen

import logging
from ui_main_window import Ui_MainWindow  # Importieren des generierten Codes
from logger import LEVEL_NAME_MAP  # LEVEL_NAME_MAP korrekt importieren
from config import DEBUG_LEVEL, MAX_FOLDER_DEPTH  # Importiere relevante Konfigurationen

logger = logging.getLogger(__name__)

class MyMainWindow(QMainWindow):
    """
    Die Klasse `MyMainWindow` repräsentiert das Hauptfenster einer GUI-Anwendung, die 
    auf PySide6 basiert. Sie enthält die Hauptlogik zur Initialisierung und Bedienung
    der Benutzeroberfläche.
    """

    def __init__(self):

        # Definiere eine interne Hilfsfunktion, um Widgets anhand ihres Namens und Typs zu finden
        def bind_widget(widget_class, name, level=2):
            """
            Bindet ein Widget (z. B. Buttons oder Komboboxen) an einen Python-Attributnamen.
            Sucht innerhalb der UI nach einem Widget mit einem spezifischen Objekt-Namen.

            :param widget_class: Klasse des Widgets (z. B. QComboBox, QPushButton).
            :param name: Der Objekt-Name des Widgets.
            :param level: Protokollierungslimit (standardmäßig 2, nur für Debugging).
            :return: Das gefundene Widget oder None, wenn kein entsprechendes Widget existiert.
            """
            widget = self.findChild(widget_class, name, Qt.FindChildrenRecursively)
            if widget and (level == 3):
                logger.debug(f"Widget gebunden: {name}") # Ausgabe erfolgt nur im Debug-Modus
            else:
                logger.warning(f"Widget NICHT gefunden: {name}")
            return widget

        # Initialisierung der QMainWindow-Basisklasse
        super().__init__()
        logger.info(f"MyMainWindow init gestartet.")

        # Lade die generierte Benutzeroberfläche aus der `Ui_MainWindow`-Klasse
        self.ui = Ui_MainWindow()  # Instanziiere die vom Designer erzeugte Klasse
        self.ui.setupUi(self)  # Wende die UI auf dieses Hauptfenster an
        logger.debug(f"Die generierte GUI wurde auf das aktuelle Fenster angewendet.")

        # Setze den Titel des Hauptfensters
        self.setWindowTitle("Outlook Email Exporter")
        logger.debug("Titelleiste des Hauptfensters initialisiert.")

        # Initialisiere die Statusleiste, um Debuginformationen anzuzeigen
        level_name = LEVEL_NAME_MAP.get(DEBUG_LEVEL, f"Unbekannt ({DEBUG_LEVEL})")
        self.statusBar().showMessage(f"Debug-Level: {level_name} | Max. Suchtiefe für Ordnerabfrage: {MAX_FOLDER_DEPTH}")
        logger.debug("Statusleiste mit Debug-Level aktualisiert.")

        # Binde Widgets aus der Benutzeroberfläche an lokale Attribute für spätere Verwendung
        # Auswahlboxen
        self.combo_postfach = bind_widget(QComboBox, "comboBox_Postfach", level=DEBUG_LEVEL)
        self.combo_verzeichnis = bind_widget(QComboBox, "comboBox_Verzeichnis", level=DEBUG_LEVEL)
        self.combo_exportziel = bind_widget(QComboBox, "comboBox_Exportziel", level=DEBUG_LEVEL)
        # Buttons
        self.button_export_msg = bind_widget(QPushButton, "pushButton_Export_MSG", level=DEBUG_LEVEL)
        self.button_export_pdf = bind_widget(QPushButton, "pushButton_Export_PDF", level=DEBUG_LEVEL)
        self.button_export_both = bind_widget(QPushButton, "pushButton_Export_Both", level=DEBUG_LEVEL)
        self.button_exit = bind_widget(QPushButton, "pushButton_Exit", level=DEBUG_LEVEL)
        # Tabellenansicht
        self.table_view = bind_widget(QTableView, "tableView_Emails", level=DEBUG_LEVEL)
        # Menüleiste
        self.action_exit = bind_widget(QObject, "actionExit", level=DEBUG_LEVEL)
        # Checkboxen
        self.checkbox_change_filename = bind_widget(QCheckBox, "checkBox_Change_Filename", level=DEBUG_LEVEL)
        self.checkbox_change_filedate= bind_widget(QCheckBox, "checkBox_Change_Filedate", level=DEBUG_LEVEL)
        self.checkbox_overwrite_file = bind_widget(QCheckBox, "checkBox_Overwrite_File", level=DEBUG_LEVEL)

        # Debugging: Listet alle in der UI definierte Widgets und ihre Namen auf
        logger.debug("Gefundene QObjects:")
        for child in self.findChildren(QObject):
            cls_name = child.metaObject().className()
            obj_name = child.objectName()
            logger.debug(f"- {cls_name}: {obj_name if obj_name else '(kein Name)'}")

        # Instanz der Dialogklasse
        self.info_dialog = None

        def show_dialog(self, dialog_text):
            if self.info_dialog is None:  # Falls der Dialog noch nicht existiert
                self.info_dialog = InfoDialog(dialog_text, self)
            self.info_dialog.show()

        def update_dialog(self, dialog_text):
            if self.info_dialog is not None:
                self.info_dialog.update_text(dialog_text)

        def close_dialog(self):
            if self.info_dialog is not None:
                self.info_dialog.close()
                self.info_dialog = None

        logger.info("MyMainWindow erfolgreich initialisiert.")


# Die erstellte Dialog-Klasse zur Steuerung des Statusdialogs
class InfoDialog(QDialog):
    def __init__(self, text, parent=None):
        super().__init__(parent)

        self.setWindowFlags(Qt.Dialog | Qt.CustomizeWindowHint | Qt.WindowTitleHint)
        layout = QVBoxLayout()
        self.label = QLabel(text)
        layout.addWidget(self.label)
        self.setLayout(layout)
        self.setWindowTitle("Info")
        self.resize(300, 100)

    def update_text(self, text):
        self.label.setText(text)
