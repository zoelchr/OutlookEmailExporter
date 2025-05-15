import os
import re
from PySide6.QtWidgets import QFileDialog
from PySide6.QtWidgets import QApplication, QMainWindow, QMessageBox, QComboBox, QTableView, QPushButton
from PySide6.QtUiTools import QUiLoader
from PySide6.QtCore import QFile
from PySide6.QtGui import QStandardItemModel, QStandardItem
from PySide6.QtCore import QObject
from PySide6.QtCore import Qt  # Importieren Sie Qt, um die CheckState-Werte zu verwenden
import win32com.client
import pythoncom

VERZEICHNIS_DATEI = "export_dirs.txt"


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

        self.combo_exportziel = self.ui.findChild(QComboBox, "comboBox_Exportziel")
        self.button_export = self.ui.findChild(QPushButton, "pushButton_Export")

        self.button_export.clicked.connect(self.exportiere_emails)
        # Exit-Button verbinden
        self.button_exit = self.ui.findChild(QPushButton, "pushButton_Exit")
        if self.button_exit:
            self.button_exit.clicked.connect(QApplication.instance().quit)
        else:
            print("Warnung: Exit-Button 'pushButton_Exit' nicht gefunden.")

        # Menüaktion 'Beenden' verbinden
        self.action_exit = self.ui.findChild(QObject, "actionExit")
        if self.action_exit:
            self.action_exit.triggered.connect(QApplication.instance().quit)
        else:
            print("Warnung: Menüaktion 'actionExit' nicht gefunden.")

        self.ordner_mapping = {}
        self.combo_postfach.currentIndexChanged.connect(self.lade_verzeichnisse)
        self.combo_verzeichnis.currentIndexChanged.connect(self.lade_emails)

        self.lade_postfaecher()

        self.lade_exportverzeichnisse()

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

    def exportiere_emails(self):
        zielpfad = self.combo_exportziel.currentText().strip()

        # Wenn der Pfad leer ist oder "..." gewählt wurde, Dialog öffnen
        if not zielpfad or zielpfad == "...":
            pfad = QFileDialog.getExistingDirectory(self.ui, "Exportverzeichnis auswählen")
            if not pfad:
                return
            zielpfad = pfad
            # ComboBox aktualisieren
            if zielpfad not in [self.combo_exportziel.itemText(i) for i in range(self.combo_exportziel.count())]:
                self.combo_exportziel.insertItem(0, zielpfad)
                self.combo_exportziel.setCurrentIndex(0)

        # Sicherstellen, dass das Verzeichnis existiert
        if not os.path.isdir(zielpfad):
            QMessageBox.warning(self.ui, "Fehler", f"Das Verzeichnis existiert nicht:\n{zielpfad}")
            return

        model = self.table_view.model()
        if model is None:
            return

        # E-Mail-Ordner aus Mapping holen
        ordner_name = self.combo_verzeichnis.currentText()
        if ordner_name not in self.ordner_mapping:
            return
        folder = self.ordner_mapping[ordner_name]

        pythoncom.CoInitialize()
        try:
            items = folder.Items
            items.Sort("[ReceivedTime]", True)

            exportiert = 0
            for i in range(min(100, model.rowCount())):
                check_item = model.item(i, 0)
                if check_item.checkState() == Qt.Checked:
                    mail = items[i]
                    dateiname = self.erstelle_dateinamen(mail)
                    pfad = os.path.join(zielpfad, dateiname)
                    mail.SaveAs(pfad, 3)  # 3 = olMSG
                    exportiert += 1

            QMessageBox.information(self.ui, "Export abgeschlossen", f"{exportiert} E-Mails wurden exportiert.")
        except Exception as e:
            QMessageBox.critical(self.ui, "Fehler beim Export", str(e))
        finally:
            self.aktualisiere_exportverzeichnis_liste(zielpfad)
            pythoncom.CoUninitialize()

    def erstelle_dateinamen(self, mail):
        betreff = mail.Subject or "ohne_Betreff"
        betreff = re.sub(r'[\\/*?:"<>|]', "_", betreff)  # ungültige Zeichen ersetzen
        datum = mail.SentOn.strftime("%Y%m%d_%H%M")
        name = f"{datum}_{betreff}.msg"
        return name[:150]  # Dateinamen nicht zu lang

    def lade_exportverzeichnisse(self):
        self.combo_exportziel.clear()
        if not os.path.exists(VERZEICHNIS_DATEI):
            return
        with open(VERZEICHNIS_DATEI, "r", encoding="utf-8") as f:
            for zeile in f:
                pfad = zeile.strip()
                if pfad:
                    self.combo_exportziel.addItem(pfad)
        self.combo_exportziel.addItem("...")  # Platzhalter für manuelle Auswahl

    def aktualisiere_exportverzeichnis_liste(self, neuer_pfad):
        pfade = []
        if os.path.exists(VERZEICHNIS_DATEI):
            with open(VERZEICHNIS_DATEI, "r", encoding="utf-8") as f:
                pfade = [zeile.strip() for zeile in f if zeile.strip()]

        # Neuen Pfad an den Anfang setzen, alte Duplikate entfernen
        if neuer_pfad in pfade:
            pfade.remove(neuer_pfad)
        pfade.insert(0, neuer_pfad)

        # Nur die letzten 20 speichern
        pfade = pfade[:20]

        with open(VERZEICHNIS_DATEI, "w", encoding="utf-8") as f:
            for pfad in pfade:
                f.write(pfad + "\n")


if __name__ == "__main__":
    app = QApplication([])
    gui = MailGUI()
    gui.show()  # statt window.show()
    app.exec()