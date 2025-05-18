from PySide6.QtWidgets import QApplication
from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QMessageBox
from logger import log
from outlook_connector import get_outlook_postfaecher

def connect_gui_signals(gui):
    """Verbindet GUI-Elemente mit den zugehörigen Funktionen."""
    log("🔗 GUI-Events werden verbunden...", level=2)

    if gui.button_exit:
        try:
            gui.button_exit.clicked.connect(on_exit_clicked)
            log("✅ Exit-Button verbunden", level=2)
        except Exception as e:
            log(f"❌ Fehler beim Verbinden des Exit-Buttons: {e}", level=0)

    if gui.action_exit:
        try:
            gui.action_exit.triggered.connect(on_exit_clicked)
            log("✅ Menü 'Exit' verbunden", level=2)
        except Exception as e:
            log(f"❌ Fehler beim Verbinden des Menü-Exit: {e}", level=0)

    # Outlook-Ladevorgang zeitverzögert ausführen (verhindert GUI-Blockade)
    QTimer.singleShot(200, lambda: load_postfaecher_async(gui))
    log("🕒 Outlook-Ladevorgang geplant (200ms Verzögerung)", level=2)

def load_postfaecher_async(gui):
    try:
        log("📥 Beginne asynchronen Outlook-Zugriff", level=2)
        postfaecher = get_outlook_postfaecher()
        if not postfaecher:
            log("⚠️ Keine Postfächer geladen – Outlook nicht verfügbar?", level=1)
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
            gui.combo_postfach.currentIndexChanged.connect(lambda index: on_postfach_changed(gui, index))
            log("📋 Postfächer nach GUI-Start geladen", level=2)
    except Exception as e:
        log(f"❌ Fehler beim asynchronen Laden der Postfächer: {e}", level=0)

def on_postfach_changed(gui, index):
    if index > 0:
        placeholder_text = "Bitte Postfach auswählen..."
        placeholder_index = gui.combo_postfach.findText(placeholder_text)
        if placeholder_index != -1:
            gui.combo_postfach.removeItem(placeholder_index)
            log("ℹ️ Platzhalter 'Bitte auswählen...' entfernt", level=2)

def on_exit_clicked():
    log("🛑 Exit ausgelöst – Anwendung wird beendet", level=1)
    QApplication.quit()