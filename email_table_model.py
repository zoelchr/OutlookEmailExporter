import logging
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex
from email_model import Email

app_logger = logging.getLogger(__name__)

class EmailTableModel(QAbstractTableModel):
    """
    Ein Qt-Modell zur Anzeige und Verwaltung von E-Mails in einer Tabelle mit Auswahl-Checkboxen.
    Jede Zeile stellt eine E-Mail dar, Spalten enthalten Metadaten (Datum, Absender, Betreff).
    """

    def __init__(self, emails: list[Email], parent=None):
        super().__init__(parent)
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        app_logger.debug(f"📦 EmailTableModel initialisiert mit {len(emails)} E-Mails")

    def rowCount(self, parent=QModelIndex()):
        return len(self.emails)

    def columnCount(self, parent=QModelIndex()):
        return 5  # Spalten: Checkbox, Datum, Absendername, Absender-E-Mail, Betreff

    def data(self, index, role=Qt.DisplayRole):
        if not index.isValid():
            return None

        row = index.row()
        col = index.column()

        if col == 0:
            if role == Qt.CheckStateRole:
                return Qt.Checked if self.checked_rows[row] else Qt.Unchecked
            if role == Qt.DisplayRole:
                return "" # <<< Das sorgt dafür, dass Qt die Checkbox korrekt anzeigt
            return None

        email = self.emails[row]

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
        if index.column() == 0:
            return Qt.ItemIsUserCheckable | Qt.ItemIsEnabled
        else:
            return Qt.ItemIsEnabled | Qt.ItemIsSelectable

    def setData(self, index, value, role=Qt.EditRole):
        if not index.isValid():
            app_logger.warning("⚠️ setData() aufgerufen mit ungültigem Index")
            return False

        row = index.row()
        col = index.column()
        app_logger.debug(f"✍️ setData(): Zeile={row}, Spalte={col}, Wert={value}, Rolle={role}")

        try:
            if col == 0 and role == Qt.CheckStateRole:
                if 0 <= row < len(self.checked_rows):
                    self.checked_rows[row] = int(value) == Qt.Checked
                    self.dataChanged.emit(index, index)
                    app_logger.debug(f"✅ Checkbox in Zeile {row} auf {'aktiviert' if value == Qt.Checked else 'deaktiviert'} gesetzt")
                    app_logger.debug(f"✍️ setData(): Zeile={row}, Spalte={col}, Wert={int(value)}, Rolle={role}")
                    return True
                else:
                    app_logger.warning(f"⚠️ Zeilenindex außerhalb des gültigen Bereichs: {row}")
        except Exception as e:
            app_logger.error(f"❌ Fehler in setData(): {e}")

        return False

    def get_selected_emails(self) -> list[Email]:
        selected = [email for i, email in enumerate(self.emails) if self.checked_rows[i]]
        app_logger.debug(f"📤 get_selected_emails(): {len(selected)} E-Mails ausgewählt")
        return selected

    def reset_data(self, emails: list[Email]):
        app_logger.debug(f"🔁 reset_data(): Modell wird mit {len(emails)} E-Mails neu befüllt")
        self.beginResetModel()
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        self.endResetModel()
        app_logger.debug("✅ Modell zurückgesetzt und aktualisiert")

    def sort(self, column, order):
        reverse = order == Qt.DescendingOrder
        app_logger.debug(f"🔃 Sortierung gestartet: Spalte={column}, Richtung={'absteigend' if reverse else 'aufsteigend'}")

        try:
            if column == 1:
                self.emails.sort(key=lambda email: email.received, reverse=reverse)
            elif column == 2:
                self.emails.sort(key=lambda email: email.sender_name.lower(), reverse=reverse)
            elif column == 3:
                self.emails.sort(key=lambda email: email.sender_email.lower(), reverse=reverse)
            elif column == 4:
                self.emails.sort(key=lambda email: email.subject.lower(), reverse=reverse)
            else:
                app_logger.debug("⏭️ Sortierung übersprungen (Checkbox-Spalte)")
                return

            self.layoutChanged.emit()
            app_logger.debug("✅ Tabellenlayout nach Sortierung aktualisiert")

        except Exception as e:
            app_logger.error(f"❌ Fehler bei Sortierung: {e}")


from PySide6.QtCore import QEvent


from PySide6.QtCore import QEvent, Qt

def editorEvent(self, event, model, option, index):
    # Überprüfe, ob der Index gültig ist und ob wir in der Checkbox-Spalte sind.
    if not index.isValid() or index.column() != 0:
        return False

    # Wir interessieren uns nur für Mausklicks (Linksklick) und Doppelklicks.
    if event.type() in (QEvent.MouseButtonRelease, QEvent.MouseButtonDblClick):
        # Verarbeite nur den linken Mausklick und prüfe, ob der Klick im aktiven Zellbereich liegt.
        if event.button() != Qt.LeftButton or not option.rect.contains(event.pos()):
            return False

        # Hier wird der aktuelle Zustand der Checkbox abgefragt.
        current_state = self.data(index, Qt.CheckStateRole)
        # Dort wird getoggelt: wenn aktuell gesetzt, dann auf Unchecked und umgekehrt.
        new_state = Qt.Unchecked if current_state == Qt.Checked else Qt.Checked

        # Setze den neuen Zustand mithilfe der setData-Methode.
        return self.setData(index, new_state, Qt.CheckStateRole)

    return False