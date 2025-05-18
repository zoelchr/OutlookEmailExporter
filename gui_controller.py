"""
gui_controller.py

Dieses Modul verknüpft die grafischen Bedienelemente (Buttons, ComboBox etc.) mit
der dahinterliegenden Logik und den Outlook-Funktionen. Es kapselt Interaktionslogik wie:

- Exit-Button und Menübindung
- zeitverzögertes Laden der Postfächer (Outlook)
- Laden und Anzeigen der Ordnerstruktur eines Postfachs
- Platzhalter-Handling in den ComboBoxen
- Benutzerhinweise bei Outlook-Problemen (QMessageBox)

Wird direkt nach GUI-Initialisierung aus `outlook_email_exporter.py` aufgerufen.
"""

from PySide6.QtWidgets import QApplication, QMessageBox, QHeaderView
from PySide6.QtCore import QTimer
from logger import log
from outlook_connector import get_outlook_postfaecher, get_outlook_ordner
from mail_reader import lade_emails
from email_table_model import EmailTableModel

def connect_gui_signals(gui):
    """Verbindet GUI-Elemente mit den zugehörigen Funktionen und initialisiert Inhalte."""
    log("🔗 GUI-Events werden verbunden...", level=2)

    # Exit-Button
    if gui.button_exit:
        try:
            gui.button_exit.clicked.connect(on_exit_clicked)
            log("✅ Exit-Button verbunden", level=2)
        except Exception as e:
            log(f"❌ Fehler beim Verbinden des Exit-Buttons: {e}", level=0)

    # Menüeintrag "Exit"
    if gui.action_exit:
        try:
            gui.action_exit.triggered.connect(on_exit_clicked)
            log("✅ Menü 'Exit' verbunden", level=2)
        except Exception as e:
            log(f"❌ Fehler beim Verbinden des Menü-Exit: {e}", level=0)

    # Verzeichnis-ComboBox deaktivieren, bis ein Postfach ausgewählt wurde
    if gui.combo_verzeichnis:
        gui.combo_verzeichnis.setEnabled(False)

    # Outlook-Postfächer asynchron laden (verhindert GUI-Blockade)
    QTimer.singleShot(200, lambda: load_postfaecher_async(gui))
    log("🕒 Outlook-Ladevorgang geplant (200ms Verzögerung)", level=2)


def load_postfaecher_async(gui):
    """Lädt Outlook-Postfächer und initialisiert die ComboBox."""
    try:
        log("📥 Beginne asynchronen Outlook-Zugriff", level=2)
        postfaecher = get_outlook_postfaecher()

        if not postfaecher:
            log("⚠️ Keine Postfächer geladen – Outlook möglicherweise nicht erreichbar", level=1)
            QMessageBox.warning(
                gui,
                "Outlook-Verbindung fehlgeschlagen",
                "Outlook konnte nicht gestartet oder verbunden werden.\n\n"
                "Bitte prüfen Sie:\n"
                "• Ist Outlook installiert?\n"
                "• Wurde Outlook korrekt beendet?\n"
                "• Läuft der Prozess 'OUTLOOK.EXE'?\n\n"
                "Versuchen Sie ggf. Outlook manuell zu starten.",
            )
            return

        if gui.combo_postfach:
            gui.combo_postfach.clear()
            gui.combo_postfach.addItem("Bitte Postfach auswählen...")
            gui.combo_postfach.addItems(postfaecher)
            gui.combo_postfach.setCurrentIndex(0)

            gui.combo_postfach.currentIndexChanged.connect(
                lambda index: on_postfach_changed(gui, index)
            )

            log("📋 Postfächer erfolgreich geladen und verbunden", level=2)

    except Exception as e:
        log(f"❌ Fehler beim Laden der Outlook-Postfächer: {e}", level=0)


def on_postfach_changed(gui, index):
    """Wird aufgerufen, wenn ein Postfach ausgewählt wurde."""
    if index > 0:
        placeholder_text = "Bitte Postfach auswählen..."
        placeholder_index = gui.combo_postfach.findText(placeholder_text)
        if placeholder_index != -1:
            gui.combo_postfach.removeItem(placeholder_index)
            log("ℹ️ Platzhalter 'Bitte Postfach auswählen...' entfernt", level=2)

        # Ordnerstruktur abrufen
        postfach_name = gui.combo_postfach.currentText()
        verzeichnisse = get_outlook_ordner(postfach_name)

        if gui.combo_verzeichnis:
            gui.combo_verzeichnis.clear()
            gui.combo_verzeichnis.addItem("Bitte Verzeichnis auswählen...")
            gui.combo_verzeichnis.addItems(verzeichnisse)
            gui.combo_verzeichnis.setEnabled(True)
            gui.combo_verzeichnis.setCurrentIndex(0)

            gui.combo_verzeichnis.currentIndexChanged.connect(
                lambda index: on_verzeichnis_changed(gui, index)
            )
            log(f"📂 {len(verzeichnisse)} Verzeichnisse für '{postfach_name}' geladen", level=2)


def on_verzeichnis_changed(gui, index):
    """Reaktion auf Verzeichniswahl: Checkbox-Platzhalter entfernen + Tabelle befüllen."""
    if index > 0:
        # Platzhalter entfernen
        placeholder_text = "Bitte Verzeichnis auswählen..."
        placeholder_index = gui.combo_verzeichnis.findText(placeholder_text)
        if placeholder_index != -1:
            gui.combo_verzeichnis.removeItem(placeholder_index)
            log("ℹ️ Platzhalter 'Bitte Verzeichnis auswählen...' entfernt", level=2)

        # E-Mail-Tabelle vorbereiten
        postfach_name = gui.combo_postfach.currentText()
        ordner_pfad = gui.combo_verzeichnis.currentText()

        # 🧠 Fehlervermeidung: Postfachname aus Pfad entfernen (falls enthalten)
        if ordner_pfad.startswith(postfach_name + "/"):
            ordner_pfad = ordner_pfad[len(postfach_name) + 1:]

        log(f"📨 Starte Mail-Import für Postfach='{postfach_name}', Ordner='{ordner_pfad}'", level=2)

        emails = lade_emails(postfach_name, ordner_pfad)
        log(f"📊 Tabelle wird mit {len(emails)} E-Mails befüllt", level=2)

        model = EmailTableModel(emails)
        gui.table_view.setModel(model)

        # Tabellenkopf vorbereiten
        header = gui.table_view.horizontalHeader()
        header.setStretchLastSection(False)

        # Resize-Strategie je Spalte
        header.setSectionResizeMode(0, QHeaderView.Fixed)  # Checkbox
        header.setSectionResizeMode(1, QHeaderView.Interactive)  # Datum
        header.setSectionResizeMode(2, QHeaderView.Interactive)  # Name
        header.setSectionResizeMode(3, QHeaderView.Interactive)  # E-Mail
        header.setSectionResizeMode(4, QHeaderView.Stretch)  # Betreff

        # Mindestbreiten setzen
        # Individuelle Mindestbreiten pro Spalte
        header.setMinimumSectionSize(10)  # kleiner Basisschutz
        gui.table_view.setColumnWidth(0, 26)  # Checkbox wirklich schmal
        gui.table_view.setColumnWidth(1, 120)
        gui.table_view.setColumnWidth(2, 180)
        gui.table_view.setColumnWidth(3, 220)

        gui.table_view.setEnabled(True)


def on_exit_clicked():
    """Beendet das Programm."""
    log("🛑 Exit ausgelöst – Anwendung wird beendet", level=1)
    QApplication.quit()