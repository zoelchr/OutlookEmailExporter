"""
gui_controller.py

Dieses Modul verknüpft die grafischen Bedienelemente (Buttons, ComboBox etc.) mit
der dahinterliegenden Logik und den Outlook-Funktionen. Es kapselt Interaktionslogik wie:

- Exit-Button und Menübindung
- zeitverzögertes Laden der Postfächer (Outlook)
- Laden und Anzeigen der Ordnerstruktur eines Postfachs
- Platzhalter-Handling in den ComboBoxen
- Benutzerhinweise bei Outlook-Problemen (QMessageBox)

### Zentrale Rolle des Parameters `gui`:
Der Parameter `gui` spielt eine entscheidende Rolle in allen hier definierten Funktionen. Er stellt die Schnittstelle zu den GUI-Elementen (z. B. Buttons, ComboBoxes, Tabellenansichten) dar und ermöglicht die Verknüpfung dieser Elemente mit der dahinterliegenden Logik. 
`gui` enthält unter anderem:
- Referenzen auf GUI-Komponenten wie `button_exit`, `combo_postfach` oder `table_view`.
- Signal-Slot-Verbindungen, die Benutzeraktionen (z. B. Klicks oder Auswahländerungen) mit den passenden Funktionen verbinden.

### Definition und Instanziierung:
Die Klasse `gui` wird typischerweise als Teil der GUI-Initialisierung definiert. Der Ursprung dieser Klasse liegt in einer externen Instanz, die aus einer separaten Datei stammt. 
In diesem Fall:
- Die `gui`-Instanz wird im übergeordneten Modul `outlook_email_exporter.py` erzeugt und in das `gui_controller.py`-Modul übergeben.

### Aufruf:
Das Modul `gui_controller.py` wird direkt nach der GUI-Initialisierung aus `outlook_email_exporter.py` aufgerufen.
"""
import logging
from PySide6.QtWidgets import QApplication, QMessageBox, QHeaderView, QAbstractItemView, QDialog, QLabel, QVBoxLayout
from PySide6.QtCore import QTimer, Qt

from outlook_connector import get_outlook_postfaecher, get_outlook_ordner, lade_emails
from email_table_model import EmailTableModel
from exportziel_manager import connect_gui_signals_exportziel_manager, get_exportziel_manager

from export_manager import ExportManager


app_logger = logging.getLogger(__name__)

placeholder_text_postfach = "Bitte Postfach auswählen..."
placeholder_text_verzeichnis = "Bitte Verzeichnis auswählen..."


def connect_gui_signals(gui):
    """Verbindet GUI-Elemente mit den zugehörigen Funktionen und initialisiert Inhalte.

    Hinweis:
    Die Aktivierung der combo_postfach erfolgt hier nicht, sondern erst in der Funktion `load_postfaecher_async`.
    Dies liegt daran, dass die Aktivierung der combo_postfach davon abhängig ist, dass eine erfolgreiche Verbindung
    zu Outlook hergestellt wurde. Dieses Vorgehen verhindert, dass die GUI-Elemente irrtümlich aktiviert werden,
    ohne dass entsprechende Daten zur Verfügung stehen.
    """
    app_logger.debug("Verbinde GUI-Signale mit Logikfunktionen...")


    # Exit-Button
    if gui.button_exit:
        try:
            gui.button_exit.clicked.connect(on_exit_clicked)
            app_logger.debug("Exit-Button verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des Exit-Buttons: {e}")

    # Menüeintrag "Exit"
    if gui.action_exit:
        try:
            gui.action_exit.triggered.connect(on_exit_clicked)
            app_logger.debug("Exit-Menüeintrag verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des Exit-Menüeintrag: {e}")

    # Verzeichnis-ComboBox zunächst deaktivieren, da erst ein Postfach ausgewählt werden muss
    if gui.combo_verzeichnis:
        gui.combo_verzeichnis.setEnabled(False)

    # Outlook-Postfächer asynchron laden (verhindert GUI-Blockade)
    # Hinweis: Die Aktivierung der combo_postfach erfolgt bewusst nicht in connect_gui_signals,
    # sondern erst in load_postfaecher_async, da sie von einer erfolgreichen Outlook-Verbindung abhängt.
    QTimer.singleShot(200, lambda: load_postfaecher_async(gui))
    app_logger.debug("Outlook-Ladevorgang geplant (200ms Verzögerung)")

    # In der GUI-Initialisierung
    connect_gui_signals_exportziel_manager(gui)
    app_logger.debug(f"Combox exportziel: {gui.combo_exportziel} mit Logik verbunden.")

    # Export-Buttons initial deaktivieren
    update_export_buttons_state(gui)


   # MSG-Export-Button
    if gui.button_export_msg:
        try:
            gui.button_export_msg.clicked.connect(lambda: on_export_msg_clicked(gui))
            app_logger.debug("MSG-Export-Button verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des MSG-Export-Buttons: {e}")


    # PDF-Export-Button
    if gui.button_export_pdf:
        try:
            gui.button_export_pdf.clicked.connect(lambda: on_export_pdf_clicked(gui))
            app_logger.debug("PDF-Export-Button verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des PDF-Export-Buttons: {e}")


    # Kombi-Export-Button (MSG + PDF)
    if gui.button_export_both:
        try:
            gui.button_export_both.clicked.connect(lambda: on_export_both_clicked(gui))
            app_logger.debug("Kombi-Export-Button verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des Kombi-Export-Buttons: {e}")


def load_postfaecher_async(gui):
    """
    Lädt die verfügbaren Outlook-Postfächer und initialisiert die zugehörige Auswahl-ComboBox in der GUI.

    Hinweis:
    Die Aktivierung der combo_postfach erfolgt bewusst nicht in der Funktion `connect_gui_signals`, 
    sondern erst hier in `load_postfaecher_async`. Hintergrund ist, dass die Aktivierung der combo_postfach 
    davon abhängt, dass eine Verbindung zu Outlook erfolgreich hergestellt werden kann. 
    Ohne eine valide Verbindung bleiben die GUI-Interaktionen mit der ComboBox deaktiviert, 
    um Fehlerszenarien oder ungültige Benutzeraktionen zu vermeiden.

    Parameter:
    gui : QObject
        Die GUI-Instanz, welche die ComboBox enthält, die mit den Postfächern gefüllt wird.

    Ablauf:
    1. Versucht, asynchron die Liste der verfügbaren Outlook-Postfächer zu laden.
    2. Falls erfolgreich:
        - Füllt die combo_postfach mit den Postfachnamen.
        - Aktiviert die combo_postfach.
        - Verbindet die Auswahländerungen (Signal-Slot-Verknüpfung).
    3. Bei Fehlermeldungen oder Ausnahmen:
        - Zeigt eine entsprechende Warnung an und deaktiviert die relevanten GUI-Aktionen.
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
    """Wird aufgerufen, wenn ein Postfach ausgewählt wurde, d.h. index > 0."""

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
    """Reaktion auf Verzeichniswahl: Checkbox-Platzhalter entfernen + Tabelle befüllen."""

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

    # Aktualisiere den Zustand der Export-Buttons
    update_export_buttons_state(gui)


def on_export_msg_clicked(gui):
    """
    Verarbeitet das Klicken auf den 'Export MSG'-Button und startet den Export
    der ausgewählten Emails als MSG-Dateien.

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

    Hinweis:
    Bei Fehlern oder ungültigen Zuständen, wie z. B. einem leeren Datenmodell
    oder keinem definierten Exportverzeichnis, wird der Ablauf abgebrochen und
    der Benutzer entsprechend informiert.
    """
    app_logger.debug("'Export MSG'-Button wurde geklickt.")

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

        # 4. Export durchführen
        email_exporter.export_emails()

        # 5. Erfolgsrückmeldung an den Benutzer
        QMessageBox.information(gui, "Export erfolgreich", "MSG-Export erfolgreich abgeschlossen!")
        app_logger.debug("MSG-Export erfolgreich durchgeführt.")

    except Exception as e:
        # Fehlerprotokollierung und Benutzerhinweis
        app_logger.error(f"Fehler beim MSG-Export: {e}")
        QMessageBox.critical(gui, "Export fehlgeschlagen", f"Ein Fehler ist aufgetreten: {e}")


def on_export_pdf_clicked(gui):
    """Aktion ausführen, wenn der PDF-Export-Button geklickt wird."""
    app_logger.debug("PDF-Export-Button wurde geklickt.")

    # try:
    #     # Tabellenmodell abrufen
    #     table_model = gui.table_view.model()
    #     if not table_model:
    #         app_logger.warning("Kein Tabellenmodell gefunden. Der Export wird abgebrochen.")
    #         QMessageBox.warning(gui, "Export fehlgeschlagen", "Keine Daten in der Tabelle verfügbar!")
    #         return
    #
    #     # Ausgewählte E-Mails basierend auf Checkboxen abrufen
    #     selected_emails = table_model.get_selected_emails()
    #
    #     if not selected_emails:
    #         app_logger.warning("Keine Zeilen ausgewählt. Der Export wird abgebrochen.")
    #         QMessageBox.warning(gui, "Keine Auswahl", "Bitte wählen Sie mindestens eine Zeile zur Verarbeitung aus!")
    #         return
    #
    #     # ExportManager initialisieren
    #     export_manager = ExportManager()
    #     export_manager.set_export_path(gui.current_export_path)  # Exportpfad setzen
    #
    #     # Exportieren der ausgewählten E-Mails
    #     for email in selected_emails:
    #         export_manager.save_as_pdf(email)
    #
    #     QMessageBox.information(gui, "Export abgeschlossen", "Die ausgewählten E-Mails wurden erfolgreich als PDF exportiert.")
    #     app_logger.debug(f"{len(selected_emails)} PDF-Dateien wurden erfolgreich gespeichert.")
    # except Exception as e:
    #     app_logger.error(f"Fehler beim Exportieren von PDFs: {e}")
    #     QMessageBox.critical(gui, "Fehler", f"Ein Fehler ist bei der Verarbeitung aufgetreten: {str(e)}")


def on_export_both_clicked(gui):
    """Aktion ausführen, wenn der Kombi-Export-Button geklickt wird."""
    app_logger.debug("Beide Exporte gestartet (MSG und PDF).")
    # Export-Logik für beide Formate hier einfügen
    QMessageBox.information(gui, "Export PDF + MSG", "Kombi-Export erfolgreich gestartet!")


def on_exit_clicked():
    """Beendet das Programm."""
    app_logger.debug("Exit ausgelöst – Anwendung wird beendet")
    QApplication.quit()


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

    :param gui: Die Instanz der GUI.
    :return: True, wenn ein valides Exportziel ausgewählt wurde, sonst False.
    """
    combo_exportziel = gui.combo_exportziel

    # Hole den aktuellen Index und die Anzahl der Elemente in der ComboBox
    current_index = combo_exportziel.currentIndex()
    total_items = combo_exportziel.count()

    # Überprüfen, ob der Index ungültig ist
    if current_index < 0 or current_index >= total_items:
        return False

    # Hole den aktuell ausgewählten Text
    selected_text = combo_exportziel.currentText()

    # Hole die ungültigen Texte direkt aus dem Exportziel-Manager
    exportziel_manager = get_exportziel_manager(gui)  # Die im ExportzielManager definierten Texte holen
    invalid_entries = [exportziel_manager.placeholder_text, exportziel_manager.new_target_text]

    # Prüfen, ob der Text ungültig ist
    if selected_text in invalid_entries:
        return False

    # Wenn alle Bedingungen erfüllt sind, ist die Auswahl valide
    return True


def show_loading_dialog(parent, message, duration=None):
    """
    Erstellt ein Dialogfenster und zeigt eine Nachricht an.

    :param parent: Das Hauptfenster der GUI (Parent für den Dialog).
    :param message: Die Nachricht, die im Dialog angezeigt wird.
    :param duration: (Optional) Zeit in Millisekunden, nach der das Dialogfenster automatisch geschlossen wird.
    """
    # Dialog erstellen
    dialog = QDialog(parent)
    dialog.setWindowFlags(Qt.FramelessWindowHint | Qt.Dialog)  # Entfernt die Titelzeile
    dialog.setModal(True)  # Blockiert Benutzerinteraktionen mit dem Hauptfenster

    # Layout und Nachricht
    layout = QVBoxLayout(dialog)
    label = QLabel(message)
    label.setAlignment(Qt.AlignCenter)
    layout.addWidget(label)

    dialog.setLayout(layout)
    dialog.resize(300, 100)  # Dialogfenstergröße definieren
    dialog.show()

    # Automatisches Schließen nach einer Zeit, falls `duration` angegeben ist
    if duration:
        QTimer.singleShot(duration, dialog.accept)

    return dialog


# def onExportTargetChanged(self, index):
#     """
#     Wird aufgerufen, wenn die Auswahl in der ComboBox für Exportziele geändert wird.
#     """
#     app_logger.debug(f"Signal onExportTargetChanged aktiviert. Index: {index}")
#     target = self.combo_exportziel.currentText()
#     app_logger.debug(f"Gewählter Eintrag: {target}")
#
#     # Wenn der ausgewählte Text der Platzhalter ist, keine Aktion ausführen
#     if target == self.placeholder_text:
#         app_logger.debug(f"Platzhaltertext ('{target}') ausgewählt, keine Änderung.")
#         return
#
#     # Wenn ein neues Ziel ausgewählt werden sollte, Dialog öffnen
#     if target == self.new_target_text:
#         app_logger.debug(f"Option 'Neues Exportziel ('{target}') wählen...' ausgewählt.")
#         self.handleNewTarget()
#         return
#
#     # Ziel validieren
#     if not self.ist_zulaessiger_pfad(target):
#         app_logger.error(f"Ungültiger Pfad ausgewählt: {target}")
#         return
#
#     app_logger.debug(f"Valider und ausgewählter Exportpfad: {target}")
#     # Speicher das neue Ziel und update history
#     self.saveExportTarget(target)