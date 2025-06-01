"""
email_table_model.py

Dieses Modul implementiert ein Qt-Tabellenmodell zur Darstellung und Verwaltung von E-Mails in einer Tabelle.
Die Tabelle enthält Informationen wie Empfangsdatum, Absender, Betreff und eine Checkbox, um E-Mails auszuwählen.

### Hauptaufgaben:
1. **Darstellung der E-Mails**:
   - Verwendung der Klasse `Email`, die die einzelnen E-Mails repräsentiert.
   - Jede Zeile der Tabelle entspricht einer E-Mail, und die Spalten enthalten Details wie:
     - Empfangsdatum und Uhrzeit
     - Name und E-Mail-Adresse des Absenders
     - Betreff der Nachricht
     - Checkbox zur Auswahl.

2. **Interaktive Tabelle**:
   - Unterstützung für Checkboxen, die es dem Benutzer ermöglichen, einzelne E-Mails für Aktionen (z. B. Export) zu markieren.
   - Spaltenbeschriftungen und datenabhängige Zell-Inhalte werden dynamisch bereitgestellt.

3. **Flexibilität und Reaktivität**:
   - Änderungen an den Daten des Modells werden umgehend in der Tabelle aktualisiert.
   - Unterstützt durch Rollen wie `Qt.CheckStateRole` (für Checkboxen) und `Qt.DisplayRole` (für anzuzeigende Daten).

### Übersicht: Methoden
1. **`__init__`**
   Initialisiert das Modell mit einer Liste von `Email`-Objekten. Erstellt zusätzlich eine Liste zum Verfolgen des Zustands der Checkboxen.

2. **`rowCount`**
   Gibt die Anzahl der Zeilen im Modell zurück, entsprechend der Anzahl der E-Mails.

3. **`columnCount`**
   Gibt die Anzahl der Spalten im Modell zurück. Die Tabelle hat feste Spalten:
   - 0: Checkbox
   - 1: Empfangsdatum
   - 2: Name des Absenders
   - 3: E-Mail-Adresse des Absenders
   - 4: Betreff.

4. **`data`**
   Stellt Daten für jede Zelle basierend auf der angeforderten Rolle bereit:
   - `Qt.DisplayRole`: Liefert die anzuzeigenden Werte (z. B. Datum, Betreff).
   - `Qt.CheckStateRole`: Liefert den Status der Checkbox (angekreuzt oder nicht).

5. **`headerData`**
   Definiert die Kopfzeilen der Tabelle. Bietet benutzerdefinierte Beschriftungen für die horizontalen Tabellenköpfe (z. B. "✓", "Betreff").

6. **`flags`**
   Legt fest, welche Aktionen für die Daten erlaubt sind (z. B. Checkbox anklicken, Daten kopieren).

7. **`setData`**
   Ermöglicht, Daten im Modell zu aktualisieren (z. B. die Checkboxen zu setzen oder zurückzusetzen).

8. **`get_selected_emails`**
   Gibt die vom Benutzer ausgewählten `Email`-Objekte zurück, die über die Checkboxen markiert wurden.

9. **`reset_data`**
   Lädt neue E-Mails in das Modell und initialisiert die Checkboxen-Liste neu.

10. **`sort`**
    Implementiert die Sortierung der E-Mails basierend auf einer bestimmten Spalte.

### Flexibilität:
- Kann leicht erweitert werden, z. B. um zusätzliche Spalten hinzuzufügen.
- Unterstützt benutzerdefinierte Rollen, wenn zusätzliche Funktionalitäten benötigt werden.

### Verknüpfung:
Wird typischerweise in Verbindung mit einer `QTableView` verwendet, die das Modell rendert und Interaktionen abbildet.
"""
import logging
from PySide6.QtCore import Qt, QAbstractTableModel, QModelIndex, QEvent
from email_model import Email

# Erstellen eines Loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

class EmailTableModel(QAbstractTableModel):
    """
    Ein Qt-Modell zur Anzeige und Verwaltung von E-Mails in einer Tabelle mit Auswahl-Checkboxen.
    Jede Zeile stellt eine E-Mail dar, Spalten enthalten Metadaten (Datum, Absender, Betreff).
    """

    def __init__(self, emails: list[Email], parent=None):
        """
        Konstruktor für die Klasse `EmailTableModel`.

        Initialisiert ein neues Tabellenmodell, welches die Darstellung und Verwaltung
        von E-Mails in einer tabellarischen Oberfläche ermöglicht. Die übergebenen Daten
        werden in der Instanz gespeichert und eine Kontrollliste (`checked_rows`) wird
        erstellt, um den Status der Checkboxen für jede Zeile zu verwalten.

        Parameter:
            emails (list[Email]): Eine Liste von `Email`-Objekten, die die Daten für 
                                  das Tabellenmodell bereitstellen.
            parent (QObject, optional): Das übergeordnete Objekt, standardmäßig `None`.

        Hauptaufgaben:
        - Speichern der E-Mails für die tabellarische Darstellung.
        - Initialisieren der Checkbox-Statusliste (`checked_rows`) mit `False` für jede Zeile
          (indiziert, dass die Checkboxen in der Tabelle standardmäßig deaktiviert sind).
        - Protokollieren der Gesamtanzahl der initialisierten E-Mails zur Debugging-Zwecken.

        Beispiel:
            emails = [
                Email("01.01.2025 12:30", "John Doe", "john.doe@example.com", "Betreff A", outlook_item),
                Email("02.01.2025 14:45", "Jane Smith", "jane.smith@example.com", "Betreff B", outlook_item)
            ]

            model = EmailTableModel(emails)
        """
        super().__init__(parent)
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        app_logger.debug(f"EmailTableModel initialisiert mit {len(emails)} E-Mails.")


    def rowCount(self, parent=None):
        """
        Bestimmt die Anzahl der Zeilen im Tabellenmodell.

        Das Modell stellt jede E-Mail als eine Zeile dar. Die Gesamtanzahl der Zeilen
        entspricht daher der Anzahl der gespeicherten E-Mails.

        Parameter:
            parent (QModelIndex, optional): Wird hier nicht verwendet, da das Modell flach ist
            (ohne Verschachtelung).

        Rückgabewert:
            int: Die Anzahl der E-Mails im Modell (entspricht der Zeilenanzahl in der Tabelle).
        """
        return len(self.emails)


    def columnCount(self, parent=None):
        """
        Gibt die Anzahl der Spalten zurück, die in der Tabelle angezeigt werden.

        Die Tabelle hat genau 5 Spalten:
        1. Checkbox (zur Auswahl der E-Mail)
        2. Datum & Uhrzeit der E-Mail
        3. Name des Absenders
        4. E-Mail-Adresse des Absenders
        5. Betreff der Nachricht

        Parameter:
            parent (QModelIndex, optional): Wird hier nicht verwendet, da das Modell flach ist
            (ohne Verschachtelung).

        Rückgabewert:
            int: Die feste Anzahl der Spalten (5).
        """
        return 5

    def data(self, index, role):
        """
        Liefert die darzustellenden oder verwendeten Daten für eine bestimmte Zelle im Modell.

        Die Methode wird von Views (wie einer QTableView) aufgerufen, um die Daten für eine bestimmte
        Kombination aus Spalte, Zeile und Rolle bereitzustellen. Die Rückgabe hängt von der Spaltenposition
        und der angeforderten Rolle ab.

        Parameter:
            index (QModelIndex): Index der Zelle, für die die Daten angefordert werden. Enthält Informationen
                             zu Zeile und Spalte.
            role (Qt.ItemDataRole): Rolle, die angibt, wie die Daten verwendet werden sollen.
                                Häufig verwendete Rollen sind:
                                - Qt.DisplayRole: Daten, die dargestellt werden sollen.
                                - Qt.CheckStateRole: Status einer Checkbox.

        Rückgabewert:
        - Für `Qt.CheckStateRole` in Spalte 0: Gibt den Status der Checkbox (`Qt.Checked` oder `Qt.Unchecked`) zurück.
        - Für `Qt.DisplayRole`: Gibt die Textdarstellung für die jeweilige Spalte zurück:
            - Spalte 1: Empfangsdatum der E-Mail.
            - Spalte 2: Name des Absenders.
            - Spalte 3: E-Mail-Adresse des Absenders.
            - Spalte 4: Betreff der E-Mail.
        - None: Wenn der Index ungültig ist oder eine nicht unterstützte Rolle angefordert wird.

        Besonderheiten:
        - Spalte 0 (Checkbox-Spalte):
            - CheckStateRole: Liefert den Zustand der Checkbox (basierend auf der `checked_rows`-Liste).
            - DisplayRole: Gibt einen leeren String zurück, um die korrekte Anzeige der Checkbox zu ermöglichen.
        - Andere Spalten (1 bis 4): Die jeweilige Eigenschaft des ausgewählten `Email`-Objekts wird basierend
          auf der Spaltennummer zurückgegeben.

        Beispiel:
            Für `role=Qt.DisplayRole` und `col=1` wird das Empfangsdatum der E-Mail zurückgegeben.

        """
        if not index.isValid():
            return None

        # Hole Zeilen- und Spalteninformationen aus dem Index
        row = index.row()
        col = index.column()
        #app_logger.debug(f"Folgendes data()-Element wird aufgerufen – Zeile {row}, Spalte {col}, Rolle {role}")
        app_logger.trace(f"Folgendes data()-Element wird aufgerufen – Zeile {row}, Spalte {col}, Rolle {role}")

        # Verarbeite die Checkbox-Spalte (Spalte 0)
        # Wenn die View den Status einer Checkbox für die erste Spalte (col == 0) abfragt, wird zurückgegeben:
        # - ob das Element angehakt ist
        # - (Qt.Checked) oder nicht (Qt.Unchecked).
        # Der Status steht in der Liste self.checked.
        #
        # Qt.CheckStateRole (Rolle 12)
        # - Diese Rolle repräsentiert den Status einer Checkbox in einer Zelle.
        # - Erwartet entweder die Konstante Qt.Checked oder Qt.Unchecked zurückzugeben, abhängig davon, ob die Checkbox aktiviert oder deaktiviert ist.
        if col == 0:
            if role == Qt.CheckStateRole:
                # Status der Checkbox: Geprüft oder ungeprüft
                app_logger.trace(f"Spalte 0, Rolle Qt.CheckStateRole: {Qt.Checked if self.checked_rows[row] else Qt.Unchecked}")
                return Qt.Checked if self.checked_rows[row] else Qt.Unchecked
            # if role == Qt.DisplayRole:
            #     # Normalerweise reicht es bei einer Checkbox einen Leerstring zurückgeben, um die Checkbox korrekt anzeigen zu lassen.
            #     # Für Debug-Text ausgeben kann es aber auch Sinn Qt.DisplayRole zu unterstützen.
            #     app_logger.debug(f"Spalte 0, Rolle Qt.DisplayRole: {str(self.checked_rows[row])}")
            #     return str(self.checked_rows[row]) # Gibt "True" oder "False" als String zurück

            # Grundsätzlich werden zwei weitere Rollen abgefragt, die aber für diesen Anwendungsfall nicht benötigt werden.
            #
            # Qt.DecorationRole (Rolle 1)
            # - Wird verwendet, um Dekorationen (z. B. Icons, Bilder) zurückzugeben.
            # - Wenn Daten für diese Rolle angefordert werden, gibt das Modell ein QIcon, ein QPixmap oder eine andere grafische Darstellung zurück.
            # - Verwendet für Zellen, die z. B. kleine Symbole oder Bilder anstelle von Text anzeigen sollen.
            #
            # Qt.EditRole (Rolle 6)
            # - Diese Rolle wird abgefragt, wenn der Benutzer eine Zelle in den Bearbeitungsmodus bringt.
            # - Wenn Qt.EditRole zurückgegeben wird, gibt das Modell den Wert zurück, der bearbeitet werden sollte (die Rohdaten).
            # - Wird auch verwendet, um zu prüfen, ob Daten in der Zelle bearbeitbar sind oder welchen Wert eine Zelle speichern soll, nachdem der Benutzer sie ändert.
            return None # Wenn die Rolle nicht Qt.DisplayRole oder Qt.CheckStateRole ist, wird None zurückgegeben.

        # Hole die E-Mail, die zur aktuellen Zeile gehört
        email = self.emails[row]

        # Verarbeitung für `Qt.DisplayRole`
        if role == Qt.DisplayRole:
            if col == 1:
                return email.received  # Empfangsdatum und -zeit
            elif col == 2:
                return email.sender_name  # Name des Absenders
            elif col == 3:
                return email.sender_email  # E-Mail-Adresse des Absenders
            elif col == 4:
                return email.subject  # Betreff der Nachricht

        # Wenn die Rolle nicht unterstützt wird, gib None zurück
        return None

    # Diese Methode gibt den Text für die Kopfzeilen zurück.
    def headerData(self, section, orientation, role):
        """
        Liefert die Kopfzeilendaten für die Tabelle (z. B. Spaltenbeschriftungen).

        Diese Methode wird von der Ansicht (View) wie QTableView aufgerufen, um die Textbeschriftung
        für die Kopfzeilen (horizontal oder vertikal) zu erhalten. Sie ermöglicht die Anpassung
        der angezeigten Titel der Tabellenköpfe basierend auf der gewünschten Orientierung
        und Rolle.

        Parameter:
            section (int): Die Position der Kopfzeile (z. B. Spalten- oder Zeilenindex).
            orientation (Qt.Orientation): Gibt die Orientierung an:
                - Qt.Horizontal: Für die Spaltenüberschriften.
                - Qt.Vertical: Für die Zeilenüberschriften (wird hier nicht explizit behandelt).
            role (Qt.ItemDataRole): Die Art der Information, die für die Kopfzeile benötigt wird.
                Häufig genutzte Rolle:
                - Qt.DisplayRole: Für den anzuzeigenden Text.

        Rückgabewert:
        str: Der Titel für die Kopfzeile, falls `role` und `orientation` zutreffen (bei Qt.DisplayRole und Qt.Horizontal).
        Andere Werte werden an die Standardmethode `super().headerData` delegiert.

        Besonderheiten:
        - Überschreibt die horizontalen Kopfzeilen mit benutzerdefinierten Titeln (z. B. Spaltennamen).
        - Nutzt eine Liste von Überschriften (`headers`), die den Spaltentiteln der Tabelle entsprechen.
        - Delegiert alle anderen Anfragen (z. B. für vertikale Ausrichtung oder Rollen, die nicht DisplayRole sind) an die Standardimplementierung.

        Beispiel:
            headers = ["✓", "Datum & Uhrzeit", "Absendername", "Absender-E-Mail", "Betreff"]
            - Spalte 0: "✓" steht für die Checkbox-Spalte.
            - Spalte 1: Zeigt Datum und Uhrzeit an.
            - Spalte 2: Name des Absenders.
            - Spalte 3: E-Mail-Adresse des Absenders.
            - Spalte 4: Betreff der E-Mail.

        """
        # Liste mit benutzerdefinierten Überschriften
        headers = ["✓", "Datum & Uhrzeit", "Absendername", "Absender-E-Mail", "Betreff"]

        # Überprüfe, ob die Rolle `Qt.DisplayRole` und die Orientierung horizontal sind
        if role == Qt.DisplayRole and orientation == Qt.Horizontal:
            # Gib die entsprechende Überschrift für die Spaltenindexnummer zurück
            return headers[section]

        # Standardrückgabe für alle anderen Fälle (z. B. vertikale Überschriften)
        return super().headerData(section, orientation, role)


    # Diese Methode gibt die Flags für eine Zelle zurück, die angeben, wie die Zelle interagiert werden kann (z.B. auswählbar, editierbar).
    def flags(self, index):
        if not index.isValid(): # Wenn der Index ungültig ist, gib Qt.NoItemFlags zurück
            return Qt.NoItemFlags
        if index.column() == 0: # Wenn die erste Spalte (Checkbox-Spalte) angefragt wird, gib Qt.ItemIsEnabled, Qt.ItemIsUserCheckable und Qt.ItemIsEditable zurück
            return Qt.ItemIsEnabled | Qt.ItemIsUserCheckable #| Qt.ItemIsEditable
        return Qt.ItemIsEnabled | Qt.ItemIsEditable | Qt.ItemIsSelectable


    #def setData(self, index, value, role=Qt.EditRole):
    def setData(self, index, value, role):
        if not index.isValid():
            app_logger.warning("setData() aufgerufen mit ungültigem Index.")
            return False

        row = index.row()
        col = index.column()
        app_logger.trace(f"setData() aufgerufen mit gültigem Index: Zeile={row}, Spalte={col}, Wert={value}, Rolle={role}")

        try:
            if col == 0 and role == Qt.CheckStateRole:
                if 0 <= row < len(self.checked_rows):
                    #self.checked_rows[row] = (value == Qt.Checked)
                    #self.dataChanged.emit(index, index)
                    current = self.checked_rows[row]
                    #print(f"Current Checked state {current}")
                    new_value = not current
                    #print(f"New Checked state {new_value}")
                    self.checked_rows[row] = new_value
                    self.dataChanged.emit(index, index, [Qt.CheckStateRole])
                    app_logger.trace(f"Checkbox in Zeile {row} auf {'aktiviert' if value == Qt.Checked else 'deaktiviert'} gesetzt")
                    app_logger.trace(f"Aufruf setData() in Zeile {row} mit Wert {int(value)} für Rolle {role}.")
                    return True
                else:
                    app_logger.warning(f"Zeilenindex außerhalb des gültigen Bereichs: {row}")
        except Exception as e:
            app_logger.error(f"Fehler in setData(): {e}")

        return False

    def get_selected_emails(self) -> list[Email]:
        selected = [email for i, email in enumerate(self.emails) if self.checked_rows[i]]
        app_logger.debug(f"📤 get_selected_emails(): {len(selected)} E-Mails ausgewählt")
        return selected

    def reset_data(self, emails: list[Email]):
        app_logger.debug(f"reset_data(): Modell wird mit {len(emails)} E-Mails neu befüllt")
        self.beginResetModel()
        self.emails = emails
        self.checked_rows = [False] * len(emails)
        self.endResetModel()
        app_logger.debug("Modell zurückgesetzt und aktualisiert")

    def sort(self, column, order):
        reverse = order == Qt.DescendingOrder
        app_logger.debug(f"Sortierung gestartet: Spalte={column}, Richtung={'absteigend' if reverse else 'aufsteigend'}")

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
                app_logger.debug("Sortierung übersprungen (Checkbox-Spalte)")
                return

            self.layoutChanged.emit()
            app_logger.debug("Tabellenlayout nach Sortierung aktualisiert")

        except Exception as e:
            app_logger.error(f"Fehler bei Sortierung: {e}")
