"""
gui_controller.py

Dieses Modul verbindet die grafischen Bedienelemente (z. B. Buttons, ComboBoxen, Tabellen) mit der zugrunde liegenden
Logik der Anwendung. Ziel ist es, eine klare Trennung zwischen der Benutzeroberfläche und der Geschäftslogik zu erreichen.

### Hauptaufgaben:
1. **Signal-Slot-Mechanismus**:
   - Verknüpft GUI-Elemente mit Funktionen, die auf Benutzereingaben reagieren (z. B. Klick auf Buttons, Auswahl in ComboBoxen).

2. **GUI-Interaktionen**:
   - Behandelt Benachrichtigungen des Benutzers (z. B. bei Fehlern oder ausstehenden Aufgaben).
   - Aktualisiert GUI-Elemente wie Tabellen, Auswahlfelder und Statusleisten basierend auf Benutzeraktionen.

3. **Exportmanagement**:
   - Koordiniert die Interaktion zwischen der GUI und dem Exportziel-Manager.
   - Validiert Benutzereingaben und Aktualisierungen in Exportziel-ComboBoxen.

4. **Outlook-Integration**:
   - Initialisiert Postfächer und lädt Ordnerstrukturen asynchron, um Benutzerinteraktionen nicht zu blockieren.
   - Zeigt informative Hinweise bei Outlook-Fehlern.

### Wichtiger Parameter:
- **`gui`**:
  Diese Instanz von `MyMainWindow` stellt die Verbindung zu den GUI-Komponenten dar.
  Beispiele:
  - `combo_postfach` (Postfachauswahl), `table_view` (E-Mail-Tabelle), `button_exit` (Beenden-Button)
  - Kombiniert Signal-Slot-Verbindungen mit den zugehörigen Logikfunktionen.

### Übersicht: Methoden
1. **`connect_gui_signals`**
   Verbindet alle relevanten GUI-Signale (z. B. Button-Klicks, ComboBox-Änderungen) mit den zugehörigen Slots (Funktionen).

2. **`load_postfaecher_async`**
   Lädt die Liste der Postfächer asynchron, um die GUI nicht zu blockieren. Zeigt währenddessen ggf. einen Ladeprozess an.

3. **`on_postfach_changed`**
   Wird ausgelöst, wenn der Benutzer die Auswahl eines Postfachs in der jeweiligen ComboBox ändert. Aktualisiert relevante Daten basierend auf der Auswahl.

4. **`on_verzeichnis_changed`**
   Reagiert auf Änderungen im Exportziel-Verzeichnis und validiert die Eingabe (z. B. ob der Pfad existiert oder erlaubt ist).

5. **`on_export_clicked`**
   Führt den Exportvorgang aus, wenn der Benutzer auf den Export-Button klickt. Überprüft vorher alle Eingaben und startet den Prozess.

6. **`update_export_buttons_state`**
   Aktiviert oder deaktiviert die Export-Buttons basierend auf dem Zustand der Benutzereingaben (z. B. ob ein gültiges Verzeichnis ausgewählt wurde).

7. **`is_valid_export_selection`**
   Validiert, ob die aktuelle Auswahl des Benutzers im Exportdialog gültig ist (z. B. kein Platzhaltertext oder ungültiger Pfad).

8. **`show_loading_dialog`**
   Zeigt ein modales Lade-Dialogfenster mit einer Nachricht an, um Benutzer über laufende Prozesse zu informieren.

9. **`on_exit_clicked`**
   Behandlung der Aktion beim Klicken auf den "Beenden"-Button. Beendet die Anwendung über eine saubere Abwicklung.

### Initialisierung und Aufruf:
- Wird in der `main.py` durch die Funktion `connect_gui_signals` nach der GUI-Initialisierung aufgerufen.
- Bindet spezifische Logikfunktionen an die benutzerdefinierte Instanz von `MyMainWindow`.

### Erweiterbarkeit:
Das Modul kann leicht für neue GUI-Funktionen erweitert werden, indem neue Widgets oder Logiken ergänzt
und mit den entsprechenden Methoden in `gui_controller.py` verknüpft werden.
"""
import logging
from PySide6.QtWidgets import QApplication, QMessageBox, QHeaderView, QAbstractItemView, QDialog, QLabel, QVBoxLayout
from PySide6.QtCore import QTimer, Qt

from outlook_connector import get_outlook_postfaecher, get_outlook_ordner, lade_emails
from email_table_model import EmailTableModel
from exportziel_manager import connect_gui_signals_exportziel_manager, get_exportziel_manager
from export_manager import ExportManager

# Erstellen eines Loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

# Platzhaltertexte für die ComboBoxen
placeholder_text_postfach = "Bitte Postfach auswählen..."
placeholder_text_verzeichnis = "Bitte Verzeichnis auswählen..."


def connect_gui_signals(gui):
    """
    Verbindet die GUI-Elemente mit zugehörigen Aktionen und Logik.

    - Verknüpft Buttons mit Funktionen (z. B. Exit-Button).
    - Initialisiert ComboBoxen und knüpft sie an asynchrone Datenverarbeitung.
    - Stellt sicher, dass Elemente erst aktiviert werden, wenn die benötigten Daten verfügbar sind.

    Argumente:
    ----------
    gui : object
        Die Haupt-GUI-Instanz mit Referenzen auf die Bedienelemente.

    Hinweis:
    Die Aktivierung der ComboBoxen (`combo_postfach`, `combo_verzeichnis`) erfolgt dynamisch nach dem Laden der Daten
    in der Funktion `load_postfaecher_async`.
    Dies liegt daran, dass die Aktivierung der combo_postfach davon abhängig ist, dass eine erfolgreiche Verbindung
    zu Outlook hergestellt wurde. Dieses Vorgehen verhindert, dass die GUI-Elemente irrtümlich aktiviert werden,
    ohne dass entsprechende Daten zur Verfügung stehen.
    """
    app_logger.debug("Verknüpfe GUI-Signale mit Logikfunktionen...")

    try:
        # Exit-Button und Menüeintrag verbinden
        if gui.button_exit:
            gui.button_exit.clicked.connect(on_exit_clicked)
            app_logger.debug("Exit-Button verbunden.")
        if gui.action_exit:
            gui.action_exit.triggered.connect(on_exit_clicked)
            app_logger.debug("Menü 'Exit' verbunden.")

        # Initial deaktiviertes Verzeichnis-Dropdown (wird später aktiviert)
        if gui.combo_verzeichnis:
            gui.combo_verzeichnis.setEnabled(False)

        # Asynchrones Laden der Outlook-Postfächer nach kleiner Verzögerung
        QTimer.singleShot(200, lambda: load_postfaecher_async(gui))
        app_logger.debug("Outlook-Ladevorgang geplant (200ms Verzögerung).")

        # Verbindung für Exportziel-Management herstellen
        connect_gui_signals_exportziel_manager(gui)
        app_logger.debug(f"Exportziel-ComboBox verbunden: {gui.combo_exportziel}.")

        # Initialzustand der Export-Buttons setzen
        update_export_buttons_state(gui)

        # Verknüpfung der Export-Buttons mit den entsprechenden Aktionen
        if gui.button_export_msg:
            gui.button_export_msg.clicked.connect(lambda: on_export_clicked(gui, "msg"))
            app_logger.debug("MSG-Export-Button verbunden.")
        if gui.button_export_pdf:
            gui.button_export_pdf.clicked.connect(lambda: on_export_clicked(gui, "pdf"))
            app_logger.debug("PDF-Export-Button verbunden.")
        if gui.button_export_both:
            gui.button_export_both.clicked.connect(lambda: on_export_clicked(gui, "both"))
            app_logger.debug("Kombi-Export-Button verbunden.")

    except Exception as e:
        app_logger.error(f"Fehler beim Verbinden der GUI-Signale: {str(e)}")


def load_postfaecher_async(gui):
    """
    Lädt asynchron die Liste der verfügbaren Outlook-Postfächer und initialisiert
    die entsprechende Auswahl-ComboBox in der GUI.

    Ablauf:
    -------
    1. Startet den Zugriff auf die Outlook-Daten, um Postfächer abzurufen.
    2. Befüllt die GUI-ComboBox für Postfächer mit den abgerufenen Daten.
    3. Aktiviert die ComboBox und verbindet deren Signal mit der entsprechenden Logik.
    4. Zeigt Fehlermeldungen in der GUI an, falls keine Verbindung möglich ist.

    Argumente:
    ----------
    gui : object
        Die GUI-Instanz, die das ComboBox-Element für die Postfächer enthält.

    Hinweis:
    Die Aktivierung der ComboBox für Postfächer erfolgt ausschließlich nach dem erfolgreichen
    Laden der Daten, um Fehlbedienungen zu vermeiden.
    """
    try:
        # Starte den Zugriff auf Outlook und lade die Liste der Postfächer.
        app_logger.debug("Beginne asynchronen Outlook-Zugriff...")
        postfaecher = get_outlook_postfaecher()

        # Export-Buttons initial deaktivieren
        update_export_buttons_state(gui)

        # Falls keine Postfächer geladen werden konnten, zeige eine Warnung und verhindere weitere Aktionen.
        if not postfaecher:
            app_logger.warning("Keine Postfächer geladen – Outlook möglicherweise nicht erreichbar.")
            QMessageBox.warning(
                gui,
                "Outlook-Verbindung fehlgeschlagen",
                "Outlook konnte nicht gestartet oder verbunden werden.\n\n"
                "Bitte prüfen Sie:\n"
                "• Ist Outlook installiert?\n"
                "• Wurde Outlook korrekt beendet?\n"
                "• Läuft der Prozess 'OUTLOOK.EXE'?\n\n"
                "Versuchen Sie ggf. Outlook manuell (neu) zu starten.",
            )
            return

        # Aktualisiere die ComboBox für die Postfächer, falls verfügbar.
        if gui.combo_postfach:
            # Entfernt alle vorhandenen Einträge in der ComboBox, um sicherzustellen,
            # dass sie vor dem Hinzufügen neuer Elemente vollständig geleert ist.
            gui.combo_postfach.clear() # ComboBox leeren...

            # Füge einen allgemeinen Auswahlhinweis und die geladenen Postfächer hinzu.
            gui.combo_postfach.addItem(placeholder_text_postfach)
            gui.combo_postfach.addItems(postfaecher)

            # Wähle standardmäßig den Hinweis-Eintrag "Bitte Postfach auswählen …" aus.
            gui.combo_postfach.setCurrentIndex(0)

            # Verbinde die Auswahländerung der ComboBox mit der entsprechenden Logik.
            # Mithilfe von `lambda index` wird die Funktion `on_postfach_changed` aufgerufen und gleichzeitig das benötigte
            # Argument `gui` übergeben. Ohne `lambda` könnte nur der Index verwendet werden, da das Signal `currentIndexChanged`
            # standardmäßig nur diesen bereitstellt.
            gui.combo_postfach.currentIndexChanged.connect(lambda index: on_postfach_changed(gui, index))

            # Aktualisiere den Zustand der Export-Buttons
            update_export_buttons_state(gui)

            # Erfolgreiche Initialisierung wird im Log dokumentiert.
            app_logger.debug("Postfächer erfolgreich geladen und verbunden.")

    except Exception as e:
        # Fehlerbehandlung: Dokumentiere unerwartete Probleme im Log und verhindere Abstürze.
        app_logger.error(f"Fehler beim Laden der Outlook-Postfächer: {e}")

        # Export-Buttons initial deaktivieren
        update_export_buttons_state(gui)


def on_postfach_changed(gui, index):
    """
    Wird ausgelöst, wenn der Benutzer ein gültiges Postfach in der entsprechenden
    ComboBox auswählt (Index > 0).

    Funktion:
    ---------
    - Entfernt einen eventuell vorhandenen Platzhalter-Eintrag aus der ComboBox.
    - Lädt und verarbeitet die Ordnerstruktur des gewählten Postfachs.
    - Aktualisiert die Tabellenansicht und initialisiert die zugehörigen GUI-Elemente.
    - Aktiviert relevante GUI-Komponenten und aktualisiert die Export-Buttons.

    Parameter:
    ----------
    gui : object
        Die GUI-Instanz, die auf die entsprechenden Bedienelemente zugreifen kann.
    index : int
        Der Index des ausgewählten Postfachs in der ComboBox.
    """

    # Aktualisiere den Zustand der Export-Buttons
    update_export_buttons_state(gui)

    if index > 0:
        # Platzhalter "Bitte Postfach auswählen …" entfernen, falls vorhanden
        # (wird nur beim ersten Aufruf benötigt)
        #placeholder_text = "Bitte Postfach auswählen..."
        placeholder_index = gui.combo_postfach.findText(placeholder_text_postfach)
        if placeholder_index != -1:
            gui.combo_postfach.removeItem(placeholder_index)
            app_logger.debug(f"Platzhalter '{placeholder_text_postfach}' entfernt")

        # Leere die Tabelle, da ein neues Postfach ausgewählt wurde.
        app_logger.debug("Leere die Tabelle, da ein neues Postfach ausgewählt wurde.")
        gui.table_view.setModel(EmailTableModel([]))  # Neues, leeres Modell setzen

        # Aktualisiere den Zustand der Export-Buttons
        update_export_buttons_state(gui)

        # Zeige dem Benutzer an, dass der Ladevorgang läuft
        # gui.statusBar().showMessage("Verzeichnisse des Postfachs werden geladen...")
        loading_dialog = show_loading_dialog(gui, "Verzeichnisse des Postfachs werden geladen... Bitte warten.")

        # Die Ordner zum gewählten Postfach werden geladen.
        postfach_name = gui.combo_postfach.currentText()
        verzeichnisse = get_outlook_ordner(postfach_name)

        # Nach dem Laden: Statusmeldung aktualisieren oder entfernen
        #gui.statusBar().showMessage(f"{len(emails)} E-Mails geladen.", 3000)  # Meldung für 5 Sekunden anzeigen
        loading_dialog.close()
        result_message = f"{len(verzeichnisse)} Verzeichnisse des Postfachs '{postfach_name}' wurden geladen."
        show_loading_dialog(gui, result_message, duration=2000)

        # Die ComboBox für Verzeichnisse wird nur angezeigt, wenn mindestens ein Verzeichnis vorhanden ist.
        if gui.combo_verzeichnis:
            # Löscht alle bestehenden Einträge in der ComboBox für Verzeichnisse,
            # um sicherzustellen, dass sie vorhinige Inhalte nicht erneut anzeigt.
            gui.combo_verzeichnis.clear()

            # Fügt einen Platzhalter-Hinweis "Bitte Verzeichnis auswählen …" zu ComboBox hinzu.
            # Dies hilft dem Benutzer, zu erkennen, dass ein Verzeichnis auszuwählen ist.
            gui.combo_verzeichnis.addItem(placeholder_text_verzeichnis)

            # Füllt die ComboBox mit den verfügbaren Verzeichnissen (Ordnern), die zuvor aus Outlook geladen wurden.
            gui.combo_verzeichnis.addItems(verzeichnisse)

            # Aktiviert die ComboBox, sodass der Benutzer mit ihr interagieren kann, nachdem sie gefüllt wurde.
            gui.combo_verzeichnis.setEnabled(True)

            # Setzt den Standardauswahlindex auf den ersten Eintrag "Bitte Verzeichnis auswählen...".
            gui.combo_verzeichnis.setCurrentIndex(0)

            # Verknüpft die GUI-Verzeichnis-ComboBox mit der Methode on_verzeichnis_changed.
            # Jedes Mal, wenn der Benutzer die Auswahl in der ComboBox ändert, wird die Funktion `on_verzeichnis_changed` aufgerufen.
            gui.combo_verzeichnis.currentIndexChanged.connect(lambda index: on_verzeichnis_changed(gui, index))

            # Protokolliert die Anzahl der geladenen Verzeichnisse in die Logging-Daten.
            app_logger.debug(f"{len(verzeichnisse)} Verzeichnisse für '{postfach_name}' geladen")

            # Aktualisiere den Zustand der Export-Buttons
            update_export_buttons_state(gui)
        else:
            # Wenn die ComboBox nicht zugreifbar war (z. B. GUI-Problem):
            # Logge eine Warnung für den Benutzer und erläutere das Problem.
            app_logger.warning(f"Kann Ordner für das Postfach {postfach_name} nicht laden – Outlook möglicherweise nicht erreichbar.")

            # Zeigt eine Warnmeldung in der GUI an, wenn Outlook-Probleme auftreten.
            # Gibt dem Benutzer auch Hinweise, um typische Probleme zu lösen (z. B. Outlook neu starten).
            QMessageBox.warning(
                gui,
                "Outlook-Verbindung fehlgeschlagen",
                "Outlook konnte nicht gestartet oder verbunden werden.\n\n"
                "Bitte prüfen Sie:\n"
                "• Ist Outlook installiert?\n"
                "• Wurde Outlook korrekt beendet?\n"
                "• Läuft der Prozess 'OUTLOOK.EXE'?\n\n"
                "Versuchen Sie ggf. Outlook manuell (neu) zu starten.",
            )

            # Aktualisiere den Zustand der Export-Buttons
            update_export_buttons_state(gui)

            # Beendet die weitere Ausführung der Methode, da keine Verzeichnisse verarbeitet werden können.
            return

    # Aktualisiere den Zustand der Export-Buttons
    update_export_buttons_state(gui)


def on_verzeichnis_changed(gui, index):
    """
    Wird ausgelöst, wenn der Benutzer ein gültiges Verzeichnis in der entsprechenden
    ComboBox auswählt (Index > 0).

    Funktion:
    ---------
    - Entfernt einen eventuell vorhandenen Platzhalter-Eintrag aus der ComboBox.
    - Lädt die E-Mails des ausgewählten Verzeichnisses und befüllt die Tabellenansicht.
    - Passt die Spaltenbreiten und Sortierungsoptionen in der Tabellenansicht an.
    - Aktiviert relevante GUI-Komponenten und aktualisiert den Zustand der Export-Buttons.

    Parameter:
    ----------
    gui : object
        Die GUI-Instanz, die auf die entsprechenden Bedienelemente zugreifen kann.
    index : int
        Der Index des ausgewählten Verzeichnisses in der ComboBox.
    """

    # Aktualisiere den Zustand der Export-Buttons
    update_export_buttons_state(gui)

    if index > 0:
        # Platzhalter "Bitte Verzeichnis auswählen …" entfernen, falls vorhanden
        # (wird nur beim ersten Aufruf benötigt)
        placeholder_index = gui.combo_verzeichnis.findText(placeholder_text_verzeichnis)
        if placeholder_index != -1:
            gui.combo_verzeichnis.removeItem(placeholder_index)
            app_logger.debug(f"Platzhalter '{placeholder_text_verzeichnis}' entfernt")

        # Holt den aktuell im Postfach-ComboBox ausgewählten Text.
        postfach_name = gui.combo_postfach.currentText()
        # Holt den aktuell im Verzeichnis-ComboBox ausgewählten Text.
        ordner_pfad = gui.combo_verzeichnis.currentText()

        # Sollte die Auswahl leer sein, wird eine Log-Warnung ausgegeben und die Methode beendet.
        if not postfach_name or not ordner_pfad:
            app_logger.warning("Kein gültiges Postfach oder Verzeichnis ausgewählt.")
            return

        # Fehlervermeidung: Postfachname aus Pfad entfernen (falls enthalten)
        if ordner_pfad.startswith(postfach_name + "/"):
            ordner_pfad = ordner_pfad[len(postfach_name) + 1:]

        # Startet den Prozess des Mail-Imports für das angegebene Postfach und den Ordner.
        # Die Funktion `lade_emails` ruft alle E-Mails aus dem Outlook-Ordner ab,
        # der durch das Postfach (`postfach_name`) und den Pfad (`ordner_pfad`) spezifiziert ist.
        # Nach Abschluss des Imports wird die Anzahl der geladenen E-Mails protokolliert.
        app_logger.debug(f"Starte Mail-Import für Postfach='{postfach_name}', Ordner='{ordner_pfad}'")

        # Zeige dem Benutzer an, dass der Ladevorgang läuft
        #gui.statusBar().showMessage("E-Mails werden geladen...")
        loading_dialog = show_loading_dialog(gui, "E-Mails werden geladen... Bitte warten.")

        # Erzeuge das Modell und setze es in die Tabelle
        emails = lade_emails(postfach_name, ordner_pfad)

        # Erzeugt ein Modell mit der Klasse `EmailTableModel`, das die Datenstruktur für E-Mails kapselt.
        model = EmailTableModel(emails)
        app_logger.debug(f"Tabelle wird mit {len(emails)} E-Mails befüllt")

        # Nach dem Laden: Statusmeldung aktualisieren oder entfernen
        #gui.statusBar().showMessage(f"{len(emails)} E-Mails geladen.", 3000)  # Meldung für 5 Sekunden anzeigen
        loading_dialog.close()
        result_message = f"{len(emails)} E-Mails wurden geladen."
        show_loading_dialog(gui, result_message, duration=2000)

        # Verknüpft das generierte Modell mit der Tabellenansicht (`table_view`) der GUI.
        # Dadurch werden die Daten aus dem Modell in der Tabellendarstellung angezeigt.
        gui.table_view.setModel(model)

        # Für die Spalten einen Mindestbreite setzen
        gui.table_view.setColumnWidth(0, 25) # Checkbox-Spalte
        gui.table_view.setColumnWidth(1, 120)
        gui.table_view.setColumnWidth(2, 180)
        gui.table_view.setColumnWidth(3, 220)

        # Resize-Strategie je Spalte
        gui.table_view.horizontalHeader().setMinimumSectionSize(25)  # kleiner Basisschutz als Mindestbreite für Spalte 0
        gui.table_view.horizontalHeader().setSectionResizeMode(0, QHeaderView.Fixed)  # Spalte mit Checkbox fixieren
        gui.table_view.horizontalHeader().setSectionResizeMode(1, QHeaderView.Interactive)  # Datum
        gui.table_view.horizontalHeader().setSectionResizeMode(2, QHeaderView.Interactive)  # Name
        gui.table_view.horizontalHeader().setSectionResizeMode(3, QHeaderView.Interactive)  # E-Mail
        gui.table_view.horizontalHeader().setSectionResizeMode(4, QHeaderView.Stretch)  # Betreff

        # Konfiguriert den horizontalen Header der Tabellenansicht (`table_view`), sodass die
        # letzte Spalte automatisch den noch verfügbaren Platz im Fenster ausfüllt.
        gui.table_view.horizontalHeader().setStretchLastSection(True)

        gui.table_view.setSortingEnabled(True)
        gui.table_view.sortByColumn(1, Qt.DescendingOrder)  # nach Datum sortieren (Spalte 1)

        gui.table_view.setEnabled(True)

        # Aktiviere Checkbox-Klickbarkeit
        gui.table_view.setEditTriggers(QAbstractItemView.AllEditTriggers)

        # Export-Buttons initial deaktivieren
        update_export_buttons_state(gui)


def on_export_clicked(gui, export_type):
    """
    Verarbeitet das Klicken auf einen der drei Export-Buttons und startet den Export
    der ausgewählten Emails in den gewünschten Formaten.
    Zusätzlich wird der Status der GUI-Checkboxen übergeben und kann in der
    Export-Logik verwendet werden.

    Ablauf:
    -------
    1. Prüfen, ob ein Datenmodell für die Tabelle verfügbar ist.
    2. Validieren, dass ein gültiges Exportverzeichnis ausgewählt wurde.
    3. Initialisieren des ExportManagers mit dem Tabellenmodell und dem Exportverzeichnis.
    4. Starten des Exports der ausgewählten E-Mails.
    5. Abschluss-Feedback an den Benutzer.

    Parameter:
    ----------
    gui : object
        Die GUI-Instanz, die Verweise auf relevante Elemente wie die Tabellenansicht
        und das Exportziel enthält.
    export_type : str
        Der Typ des Exports: "msg", "pdf" oder "both" (beide Formate).

    Hinweis:
    Der Status der Checkboxen wird zuverlässig aus der GUI gelesen und an den
    Exportprozess übergeben.
    Bei Fehlern oder ungültigen Zuständen, wie z. B. einem leeren Datenmodell
    oder keinem definierten Exportverzeichnis, wird der Ablauf abgebrochen und
    der Benutzer entsprechend informiert.
    """
    app_logger.debug(f"Export ausgelöst: Typ={export_type}")

    # 0. Status der Checkboxen abrufen
    change_filename = gui.checkbox_change_filename.isChecked()
    change_filedate = gui.checkbox_change_filedate.isChecked()
    overwrite_file = gui.checkbox_overwrite_file.isChecked()

    # 1. Prüfen, ob ein Tabellenmodell vorhanden ist
    table_model = gui.table_view.model()
    if not table_model:
        app_logger.error("Kein Tabellenmodell verfügbar. Der Export wird abgebrochen.")
        QMessageBox.warning(gui, "Export fehlgeschlagen", "Keine Daten in der Tabelle verfügbar!")
        return

    # 2. Überprüfen, ob ein gültiges Exportverzeichnis ausgewählt wurde
    export_directory = gui.combo_exportziel.currentText()
    if not export_directory:
        app_logger.error("Kein gültiges Exportverzeichnis ausgewählt. Der Export wird abgebrochen.")
        QMessageBox.warning(gui, "Export fehlgeschlagen", "Kein gültiges Exportverzeichnis ausgewählt!")
        return

    try:
        # 3. ExportManager-Instanz erstellen
        email_exporter = ExportManager(table_model, export_directory)

        # 4. Export initiieren
        if export_type in ["msg", "pdf", "both"]:
            exported_emails = email_exporter.export_emails(export_type, change_filename, change_filedate, overwrite_file)

            QMessageBox.information(gui, "Export erfolgreich", f"Anzahl exportierter Emails: {exported_emails} \nZielverzeichnis: {export_directory}")
            app_logger.debug(
                f"Export erfolgreich. "
                f"Details: export_type={export_type}, change_filename={change_filename}, change_filedate={change_filedate}, overwrite_file={overwrite_file}. "
                f"Anzahl exportierter Emails: {exported_emails}. Zielverzeichnis: {export_directory}"
            )

        else:
            QMessageBox.warning(gui, "Ungültiger Exporttyp", f"Der Exporttyp '{export_type}' ist nicht bekannt.")
            app_logger.error(f"Export abgebrochen: Unbekannter Exporttyp '{export_type}'.")

    except Exception as e:
        # Fehlerprotokollierung und Benutzerhinweis
        app_logger.error(f"Fehler beim Export ({export_type}): {e}")
        QMessageBox.critical(gui, "Export fehlgeschlagen", f"Ein Fehler ist aufgetreten: {e}")


def update_export_buttons_state(gui):
    """
    Aktualisiert den Aktivierungszustand der Export-Buttons basierend auf
    den Benutzerauswahlen in der GUI.

    Diese Funktion prüft, ob die erforderlichen Kriterien für den Export
    erfüllt sind:
    1. Ein gültiges Postfach ist ausgewählt.
    2. Ein gültiges Verzeichnis ist ausgewählt.
    3. Ein gültiges Exportziel ist definiert.

    Wenn eine der Bedingungen nicht erfüllt ist, werden die Export-Buttons
    deaktiviert, um eine falsche Verarbeitung zu verhindern.

    Parameter:
    ----------
    gui : object
        Die GUI-Instanz, die Referenzen zu den relevanten Elementen
        (ComboBoxen und Buttons) enthält.
    """
    app_logger.debug("Update des Zustands der Export-Buttons gestartet...")

    try:
        # Überprüfen, ob ein gültiges Postfach ausgewählt wurde
        is_postfach_selected = (
            gui.combo_postfach.currentText() != placeholder_text_postfach and gui.combo_postfach.currentIndex() >= 0
        )
        app_logger.debug(
            f"Postfach-Auswahl überprüft: "
            f"'{gui.combo_postfach.currentText()}' (Index: {gui.combo_postfach.currentIndex()}), "
            f"gültig = {is_postfach_selected}"
        )

        # Überprüfen, ob ein gültiges Verzeichnis ausgewählt wurde
        is_verzeichnis_selected = (
            gui.combo_verzeichnis.currentText() != placeholder_text_verzeichnis and
            gui.combo_verzeichnis.currentIndex() >= 0
        )
        app_logger.debug(
            f"Verzeichnis-Auswahl überprüft: "
            f"'{gui.combo_verzeichnis.currentText()}' (Index: {gui.combo_verzeichnis.currentIndex()}), "
            f"gültig = {is_verzeichnis_selected}"
        )

        # Überprüfen, ob ein gültiges Exportziel gewählt wurde
        is_exportziel_selected = is_valid_export_selection(gui)

        # Bestimme, ob die Buttons aktiviert werden sollten
        should_enable_buttons = (
            is_postfach_selected and
            is_verzeichnis_selected and
            is_exportziel_selected
        )

        # Setze den Aktivierungszustand der Export-Buttons
        gui.button_export_msg.setEnabled(should_enable_buttons)
        gui.button_export_pdf.setEnabled(should_enable_buttons)
        gui.button_export_both.setEnabled(should_enable_buttons)

        app_logger.debug(
            f"Export-Buttons aktualisiert: "
            f"Postfach ausgewählt = {is_postfach_selected}, "
            f"Verzeichnis ausgewählt = {is_verzeichnis_selected}, "
            f"Exportziel gültig = {is_exportziel_selected}, "
            f"Buttons aktiviert = {should_enable_buttons}"
        )
    except Exception as e:
        app_logger.error(f"Fehler beim Aktualisieren des Button-Zustands: {e}")


def is_valid_export_selection(gui) -> bool:
    """
    Überprüft, ob die aktuelle Auswahl in der ComboBox ein gültiges Exportziel darstellt.

    Diese Funktion validiert die Benutzer-Auswahl in der Export-Ziel-ComboBox (z. B. für Dateispeicherorte).
    Es wird überprüft, ob die Auswahl gültig ist, indem bestimmte ungültige Zustände wie Platzhaltertexte
    oder leere Einträge ausgeschlossen werden.

    Parameter:
        gui (MyMainWindow): Die Instanz des Hauptfensters der GUI, das die ComboBox `combo_exportziel` enthält.

    Rückgabewert:
        bool: Gibt `True` zurück, wenn die aktuelle Auswahl in der ComboBox ein gültiges Exportziel ist.
              Gibt `False` zurück, wenn die Auswahl ungültig ist (z. B. ein Platzhaltertext, keine Auswahl).

    Ablauf:
        1. Prüft, ob die Auswahl in der ComboBox innerhalb der gültigen Indizes liegt.
        2. Validiert, ob der Text der Auswahl einer ungültigen Option zugeordnet ist (z. B. Platzhalter).

    Hinweis:
    Hier erfolgt noch keine Prüfung, ob das ausgewählte Exportziel auch tatsächlich existiert oder erreichbar ist.
    """
    combo_exportziel = gui.combo_exportziel

    # Index und Anzahl der Einträge in der ComboBox abrufen
    current_index = combo_exportziel.currentIndex()
    total_items = combo_exportziel.count()

    # Ungültigen Index überprüfen
    if current_index < 0 or current_index >= total_items:
        return False

    # Text der aktuellen Auswahl abrufen
    selected_text = combo_exportziel.currentText()

    # Ungültige Texte aus dem ExportzielManager abrufen
    exportziel_manager = get_exportziel_manager(gui)  # Instanziiere den Exportzielmanager (als Singleton)
    invalid_entries = [exportziel_manager.placeholder_text, exportziel_manager.new_target_text]

    # Prüfen, ob die aktuelle Auswahl ungültig ist
    if selected_text in invalid_entries:
        return False

    # Auswahl ist valide
    return True


def show_loading_dialog(parent, message, duration=None):
    """
    Zeigt ein Lade-Dialogfenster mit einer Nachricht an.

    Dieser Dialog wird verwendet, um den Benutzer über einen Vorgang zu informieren, der eine Wartezeit erfordert,
    wie z. B. den Export oder das Laden von Daten. Der Dialog ist modales Fenster und blockiert
    alle interaktiven Aktionen mit dem Hauptfenster.

    Parameter:
        parent (QWidget): Das Hauptfenster oder Elternfenster der GUI, das dem Dialog als Parent dient.
        message (str): Der Informationstext, der im Dialog angezeigt werden soll.
        duration (int, optional): Zeit in Millisekunden, nach der der Dialog automatisch geschlossen wird.
                                  Standard ist `None`, falls keiner angegeben ist.

    Rückgabewert:
        QDialog: Die Instanz des erstellten Dialogs.
    """
    # Dialog initialisieren
    dialog = QDialog(parent)
    dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)  # Entfernt die Fensterdekorationen
    dialog.setModal(True)  # Blockiert Benutzerinteraktion im Hauptfenster

    # Layout und Label für die Nachricht erstellen
    layout = QVBoxLayout(dialog)  # Verwendet ein vertikales Layout für das Dialogfenster
    label = QLabel(message)  # Label mit der Nachricht initialisieren
    label.setAlignment(Qt.AlignCenter)  # Zentriere den Text im Dialog
    layout.addWidget(label)  # Füge das Label zum Layout hinzu

    dialog.setLayout(layout)  # Setze das Layout für den Dialog
    dialog.resize(500, 100)  # Setze die Größe des Dialogfensters
    dialog.show()  # Zeige den Dialog an

    # Falls eine Dauer angegeben ist, schließe den Dialog automatisch nach Ablauf dieser Zeit
    if duration:
        QTimer.singleShot(duration, dialog.accept)

    return dialog


def on_exit_clicked():
    """Beendet das Programm."""
    app_logger.debug("Exit ausgelöst – Anwendung wird beendet")
    QApplication.quit()