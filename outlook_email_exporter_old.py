import os
import logging
import re
import pythoncom
import win32com.client
from datetime import datetime
from dataclasses import dataclass
from typing import Optional, List, Any

from PySide6.QtWidgets import QFileDialog, QApplication, QMainWindow, QMessageBox, QComboBox, QTableView, QPushButton, QProgressDialog
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile, Qt, QAbstractTableModel, QModelIndex
from PySide6.QtCore import QObject

VERZEICHNIS_DATEI = "export_dirs.txt"

@dataclass
class OutlookEmail:
    entry_id: str
    sender: str
    date: Optional[datetime]
    subject: str
    recipient: Optional[str] = None
    attachment_count: Optional[int] = None
    size_kb: Optional[int] = None
    original_com_object: Optional[Any] = None

class EmailTableModel(QAbstractTableModel):
    def __init__(self, emails: List[OutlookEmail]):
        super().__init__()
        self.emails = emails
        self.headers = ["Auswählen", "Absender", "Datum", "Betreff"]

    def rowCount(self, parent=QModelIndex()):
        return len(self.emails)

    def columnCount(self, parent=QModelIndex()):
        return len(self.headers)

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        email = self.emails[index.row()]
        col = index.column()

        if role == Qt.DisplayRole:
            if col == 1:
                return email.sender
            elif col == 2:
                return email.date.strftime("%d.%m.%Y %H:%M") if email.date else "-"
            elif col == 3:
                return email.subject
        elif role == Qt.CheckStateRole and col == 0:
            return Qt.Unchecked

        return None

    def headerData(self, section, orientation, role):
        if orientation == Qt.Horizontal and role == Qt.DisplayRole:
            return self.headers[section]
        return None

logging.basicConfig(filename="email_export.log", level=logging.INFO,
                    format="%(asctime)s - %(levelname)s - %(message)s")

class MailGUI:
    def __init__(self):
        print("🔧 Initialisiere MailGUI...")
        loader = QUiLoader()
        ui_file = QFile("./outlook_email_exporter.ui")
        absolute_path = os.path.abspath("outlook_ordner_auflisten.ui")
        print(f"Absoluter Pfad der UI-Datei: {absolute_path}")

        if not ui_file.exists():
            print("❌ Die UI-Datei wurde nicht gefunden.")
        else:
            print("✅ Die UI-Datei wurde gefunden.")

        ui_file.open(QFile.ReadOnly)
        self.ui = loader.load(ui_file, None)
        ui_file.close()

        if self.ui is None:
            print("❌ UI konnte nicht geladen werden.")
        else:
            print("✅ UI wurde erfolgreich geladen (im Konstruktor).")

        self.combo_postfach = self.ui.findChild(QComboBox, "comboBox_Postfach")
        self.combo_verzeichnis = self.ui.findChild(QComboBox, "comboBox_Verzeichnis")
        self.table_view = self.ui.findChild(QTableView, "tableView_Emails")
        self.combo_exportziel = self.ui.findChild(QComboBox, "comboBox_Exportziel")
        self.button_export_msg = self.ui.findChild(QPushButton, "pushButton_Export_MSG")
        self.button_export_pdf = self.ui.findChild(QPushButton, "pushButton_Export_PDF")
        self.button_exit = self.ui.findChild(QPushButton, "pushButton_Exit")

        print("🧪 Widget-Status:")
        print(" - combo_postfach:", "✅" if self.combo_postfach else "❌")
        print(" - combo_verzeichnis:", "✅" if self.combo_verzeichnis else "❌")
        print(" - table_view:", "✅" if self.table_view else "❌")
        print(" - combo_exportziel:", "✅" if self.combo_exportziel else "❌")
        print(" - button_export_msg:", "✅" if self.button_export_msg else "❌")
        print(" - button_export_pdf:", "✅" if self.button_export_pdf else "❌")
        print(" - button_exit:", "✅" if self.button_exit else "❌")

        if self.button_exit:
            self.button_exit.clicked.connect(QApplication.instance().quit)
        #if self.button_export_msg:
            #self.button_export_msg.clicked.connect(self.exportiere_emails)
        #if self.button_export_pdf:
            #self.button_export_pdf.clicked.connect(self.exportiere_pdfs)

        self.ordner_mapping = {}
        self.email_liste: List[OutlookEmail] = []

        self.combo_postfach.currentIndexChanged.connect(self.lade_verzeichnisse)
        #self.combo_verzeichnis.currentIndexChanged.connect(self.lade_emails)

        #self.lade_postfaecher()
        #self.lade_exportverzeichnisse()

if __name__ == "__main__":
    print("🚀 Programmstart...")
    try:
        app = QApplication([])
        print("✅ QApplication initialisiert.")

        gui = MailGUI()
        print("✅ MailGUI wurde erstellt.")

        if gui.ui is None:
            print("❌ GUI konnte nicht geladen werden.")
        else:
            print("✅ GUI wird angezeigt...")
            gui.ui.resize(800, 600)
            gui.ui.move(100, 100)
            gui.ui.raise_()
            gui.ui.activateWindow()
            gui.ui.show()

        app.exec()
        print("✅ Event-Schleife wurde beendet.")

    except Exception as e:
        import traceback
        print("❌ Ausnahme beim Start des Programms:")
        traceback.print_exc()
