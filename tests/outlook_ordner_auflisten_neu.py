from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QComboBox, QTableView
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtCore import Qt  # Importieren Sie Qt, um die CheckState-Werte zu verwenden
import win32com.client
import pythoncom

class MailGUI:
    def __init__(self):
        loader = QUiLoader()
        ui_file = QFile("outlook_ordner_auflisten.ui")
        ui_file.open(QFile.ReadOnly)
        self.ui = loader.load(ui_file, None)  # <== wichtig: kein Parent (None)
        ui_file.close()

        # Widgets finden
        self.combo_postfach = self.ui.findChild(QComboBox, "comboBox_Postfach")
        self.combo_verzeichnis = self.ui.findChild(QComboBox, "comboBox_Verzeichnis")
        self.table_view = self.ui.findChild(QTableView, "tableView_Emails")

        self.ordner_mapping = {}
        self.combo_postfach.currentIndexChanged.connect(self.lade_verzeichnisse)
        self.combo_verzeichnis.currentIndexChanged.connect(self.lade_emails)

        self.lade_postfaecher()

    def show(self):
        self.ui.show()

    def lade_postfaecher(self):
        self.combo_postfach.clear()
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            for i in range(ns.Folders.Count):
                name = ns.Folders.Item(i + 1).Name
                self.combo_postfach.addItem(name)
        finally:
            pythoncom.CoUninitialize()

    def lade_verzeichnisse(self):
        self.combo_verzeichnis.clear()
        self.ordner_mapping.clear()

        postfachname = self.combo_postfach.currentText()
        if not postfachname:
            return

        # Liste der unerwünschten Ordnernamen
        unerwuenschte_ordner = ["Synchronisierungsprobleme", "Kontakte", "Aufgaben", "Kalender", "PersonMetadata"]

        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            ns = outlook.GetNamespace("MAPI")
            for i in range(ns.Folders.Count):
                postfach = ns.Folders.Item(i + 1)
                if postfach.Name == postfachname:
                    for ordner in postfach.Folders:
                        # Überprüfen, ob der Ordner E-Mails enthalten kann und nicht in der unerwünschten Liste ist
                        if ordner.DefaultItemType == 0 and ordner.Name not in unerwuenschte_ordner:
                            if self.hat_email_elemente(ordner):
                                self.combo_verzeichnis.addItem(ordner.Name)
                                self.ordner_mapping[ordner.Name] = ordner
                                for sub in ordner.Folders:
                                    if sub.DefaultItemType == 0 and sub.Name not in unerwuenschte_ordner and self.hat_email_elemente(sub):
                                        pfad = f"{ordner.Name}/{sub.Name}"
                                        self.combo_verzeichnis.addItem(pfad)
                                        self.ordner_mapping[pfad] = sub
                    break
        finally:
            pythoncom.CoUninitialize()

    def hat_email_elemente(self, ordner):
        """Überprüft, ob der Ordner E-Mail-Elemente enthält."""
        try:
            for item in ordner.Items:
                if item.Class == 43:  # 43 steht für MailItem
                    return True
        except Exception as e:
            print(f"Fehler beim Überprüfen des Ordners: {e}")
        return False

    def lade_emails(self):
        ordner_name = self.combo_verzeichnis.currentText()
        if not ordner_name or ordner_name not in self.ordner_mapping:
            return

        folder = self.ordner_mapping[ordner_name]

        pythoncom.CoInitialize()
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)

            model = QStandardItemModel()
            model.setHorizontalHeaderLabels(["Auswählen", "Absender", "Datum", "Betreff"])

            for i, item in enumerate(items):
                if i >= 100:
                    break
                betreff = item.Subject or "(Ohne Betreff)"
                sender = item.SenderName or "(Unbekannt)"
                datum = item.SentOn.strftime("%d.%m.%Y %H:%M") if item.SentOn else "-"

                # Checkbox hinzufügen
                checkbox_item = QStandardItem()
                checkbox_item.setCheckable(True)  # Checkbox aktivieren
                checkbox_item.setCheckState(Qt.Unchecked)  # Standardmäßig nicht ausgewählt

                model.appendRow([
                    checkbox_item,
                    QStandardItem(sender),
                    QStandardItem(datum),
                    QStandardItem(betreff)
                ])

            self.table_view.setModel(model)
            self.table_view.resizeColumnsToContents()
        except Exception as e:
            print(f"Fehler beim Laden der E-Mails: {e}")
        finally:
            pythoncom.CoUninitialize()

if __name__ == "__main__":
    app = QApplication([])
    gui = MailGUI()
    gui.show()  # statt window.show()
    app.exec()