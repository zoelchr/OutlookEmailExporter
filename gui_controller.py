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
from PySide6.QtWidgets import QApplication, QMessageBox, QHeaderView, QAbstractItemView
from PySide6.QtCore import QTimer, Qt

from outlook_connector import get_outlook_postfaecher, get_outlook_ordner, lade_emails
from email_table_model import EmailTableModel
from exportziel_manager import connect_gui_signals_exportziel_manager

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

    # PDF-Export-Button
    if gui.button_export_pdf:
        try:
            gui.button_export_pdf.clicked.connect(on_exit_clicked)
            app_logger.debug("PDF-Export-Button verbunden.")
        except Exception as e:
            app_logger.error(f"Fehler beim Verbinden des PDF-Export-Buttons: {e}")

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

        # Die Ordner zum gewählten Postfach werden geladen.
        postfach_name = gui.combo_postfach.currentText()
        verzeichnisse = get_outlook_ordner(postfach_name)

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
        gui.statusBar().showMessage("E-Mails werden geladen...")


        # Erzeuge das Modell und setze es in die Tabelle
        emails = lade_emails(postfach_name, ordner_pfad)

        # Erzeugt ein Modell mit der Klasse `EmailTableModel`, das die Datenstruktur für E-Mails kapselt.
        model = EmailTableModel(emails)
        app_logger.debug(f"Tabelle wird mit {len(emails)} E-Mails befüllt")

        # Nach dem Laden: Statusmeldung aktualisieren oder entfernen
        gui.statusBar().showMessage(f"{len(emails)} E-Mails geladen.", 3000)  # Meldung für 5 Sekunden anzeigen

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

def on_exit_clicked():
    """Beendet das Programm."""
    app_logger.debug("Exit ausgelöst – Anwendung wird beendet")
    QApplication.quit()


def update_export_buttons_state(gui):
    """
    Aktiviert oder deaktiviert die Export-Buttons basierend darauf, ob ein gültiges Postfach und ein gültiges Verzeichnis ausgewählt wurden.
    Invalid sind Platzhaltertexte oder nicht gewählte Einträge.
    """
    app_logger.debug(f"Export-Buttons aktualisieren...")

    # Platzhaltertexte für Postfach und Verzeichnis (anpassen, falls in der UI andere Texte genutzt werden)
    #invalid_postfach_text = "Bitte Postfach auswählen..."
    #invalid_verzeichnis_text = "Bitte Verzeichnis auswählen..."

    # Überprüfen, ob ein gültiges Postfach ausgewählt wurde
    if gui.combo_postfach.currentText() != placeholder_text_postfach:
        if gui.combo_postfach.currentIndex() >= 0:
            is_postfach_selected = True
        else:
            is_postfach_selected = False
    else:
        is_postfach_selected = False

    app_logger.debug(f"Ausgewähltes Postfach = '{gui.combo_postfach.currentText()}' (Index {gui.combo_postfach.currentIndex()}) und Anzeige Platzhalter-Text = {gui.combo_postfach.currentText() != placeholder_text_postfach}")

    # Überprüfen, ob ein gültiges Verzeichnis ausgewählt wurde
    is_verzeichnis_selected = (
            gui.combo_verzeichnis.currentIndex() > 0 and
            (gui.combo_verzeichnis.currentText() != placeholder_text_verzeichnis)
    )
    if gui.combo_verzeichnis.currentText() != placeholder_text_verzeichnis:
        if gui.combo_verzeichnis.currentIndex() >= 0:
            is_verzeichnis_selected = True
        else:
            is_verzeichnis_selected = False
    else:
        is_verzeichnis_selected = False

    app_logger.debug(f"Ausgewähltes Verzeichnis = '{gui.combo_verzeichnis.currentText()}' (Index {gui.combo_verzeichnis.currentIndex()}) und Anzeige Platzhalter-Text = {gui.combo_verzeichnis.currentText() != placeholder_text_verzeichnis}")

    # Buttons deaktivieren, wenn eines der Kriterien nicht erfüllt ist
    should_enable_buttons = is_postfach_selected and is_verzeichnis_selected
    gui.button_export_msg.setEnabled(should_enable_buttons)
    gui.button_export_pdf.setEnabled(should_enable_buttons)
    gui.button_export_both.setEnabled(should_enable_buttons)

    app_logger.debug(f"Export-Buttons aktualisiert: Postfach ausgewählt = {is_postfach_selected}, Verzeichnis ausgewählt = {is_verzeichnis_selected}, Buttons aktiviert = {should_enable_buttons}")

