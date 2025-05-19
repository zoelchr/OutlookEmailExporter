import logging
app_logger = logging.getLogger(__name__)
"""
email_table_model.py

Dieses Modul stellt ein Tabellenmodell auf Basis von QAbstractTableModel bereit,
das E-Mail-Daten inklusive Auswahl-Checkboxen verwaltet. Es erlaubt die Anzeige
mehrerer E-Mails mit folgenden Spalten:

- Auswahl (Checkbox)
- Datum und Uhrzeit
- Absendername
- Absender-E-Mail
- Betreff

Die E-Mail-Daten stammen aus einer Liste von `Email`-Datensätzen (siehe `email_model.py`).
"""
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from email_model import Email

class EmailTableModel(QAbstractTableModel):
    def __init__(self, emails: list[Email], parent=None):
        super().__init__(parent)
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        app_logger.debug(f"{len(emails)} E-Mails im Modell initialisiert")

    def rowCount(self, parent=QModelIndex()):
        return len(self.emails)

    def columnCount(self, parent=QModelIndex()):
        return 5  # Checkbox + 4 Daten-Spalten

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()
        email = self.emails[row]

        # Checkboxen in Spalte 0
        if col == 0:
            if role == Qt.CheckStateRole:
                return Qt.Checked if self.checked_rows[row] else Qt.Unchecked
            return None

        if role == Qt.DisplayRole:
            if col == 1:
                return email.received
            elif col == 2:
                return email.sender_name
            elif col == 3:
                return email.sender_email
            elif col == 4:
                return email.subject

        return None

    def headerData(self, section, orientation, role=Qt.DisplayRole):
        if role != Qt.DisplayRole or orientation != Qt.Horizontal:
            return None

        headers = ["✓", "Datum & Uhrzeit", "Absendername", "Absender-E-Mail", "Betreff"]
        return headers[section] if section < len(headers) else None

    def flags(self, index):
        if not index.isValid():
            return Qt.ItemIsEnabled

        base_flags = super().flags(index)
        if index.column() == 0:
            return base_flags | Qt.ItemIsUserCheckable | Qt.ItemIsEnabled
        else:
            return base_flags | Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid():
            return False

        row = index.row()
        col = index.column()
        app_logger.debug(f"setData(): Zeile={row}, Wert={value}, Rolle={role}")

        if col == 0 and role == Qt.CheckStateRole:
            self.checked_rows[row] = value == Qt.Checked
            self.dataChanged.emit(index, index)
            return True

        return False

    def get_selected_emails(self):
        selected = [email for i, email in enumerate(self.emails) if self.checked_rows[i]]
        app_logger.debug(f"{len(selected)} E-Mails ausgewählt")
        return selected

    def reset_data(self, emails: list[Email]):
        app_logger.debug(f"reset_data(): {len(emails)} E-Mails werden gesetzt")
        self.beginResetModel()
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        self.endResetModel()
        app_logger.debug("Modell erfolgreich zurückgesetzt")

    def sort(self, column, order):
        reverse = order == Qt.DescendingOrder
        app_logger.debug(f"Sortierung ausgelöst: Spalte={column}, Richtung={'absteigend' if reverse else 'aufsteigend'}")

        if column == 1:
            self.emails.sort(key=lambda email: email.received, reverse=reverse)
        elif column == 2:
            self.emails.sort(key=lambda email: email.sender_name.lower(), reverse=reverse)
        elif column == 3:
            self.emails.sort(key=lambda email: email.sender_email.lower(), reverse=reverse)
        elif column == 4:
            self.emails.sort(key=lambda email: email.subject.lower(), reverse=reverse)
        else:
            app_logger.debug("Sortierung übersprungen (Checkbox-Spalte)")
            return

        self.layoutChanged.emit()
        app_logger.debug("Tabellenlayout nach Sortierung aktualisiert")
