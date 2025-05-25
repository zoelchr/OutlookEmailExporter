import os
import logging
from PySide6.QtWidgets import QComboBox, QFileDialog, QDialog
from dotenv import load_dotenv

app_logger = logging.getLogger(__name__)

class ExportzielManager:
    def __init__(self, gui):
        """
        Initialisiert die Logik zur Behandlung der Exportziele.
        """
        self.gui = gui
        self.combo_exportziel = self.gui.combo_exportziel  # Geändertes Widget: combo_exportziel
        self.saved_export_file = "export_history.txt"  # Datei zur Speicherung der letzten Exportziele
        self.max_saved_targets = 50  # Maximalanzahl gespeicherter Exportziele
        self.max_env_targets = 5     # Maximalanzahl aus der .env-Datei
        self.max_combobox_items = 10 # Maximalanzahl von Einträgen in der ComboBox

        # Platzhalter- und zusätzlicher Text für ComboBox-Einträge
        self.placeholder_text = "Bitte vor Start des Exports ein Zielverzeichnis wählen..."
        self.new_target_text = "Neues Exportziel wählen..."

        # Verzeichnisse aus der .env-Datei laden
        load_dotenv()
        self.prefixed_targets = [os.getenv(f"EXPORT_TARGET_{i}") for i in range(1, self.max_env_targets + 1)]
        self.prefixed_targets = [t for t in self.prefixed_targets if t]  # Vorhandene (nicht-None) Ziele behalten

        # Initialisierung der ComboBox
        self.initExportTargets()


    def initExportTargets(self):
        """
        Initialisiert die Exportziel-ComboBox mit den Einträgen aus der .env-Datei
        und der Liste zuletzt verwendeter Verzeichnisse.
        Vermeidet Duplikate zwischen .env und der export_history.txt.
        """
        self.combo_exportziel.clear()
        self.combo_exportziel.addItem(self.placeholder_text)

        # Setze eine normalisierte Liste der Ziele aus der .env-Datei
        normalized_env_targets = [os.path.normpath(target).replace("\\", "/") for target in self.prefixed_targets]

        # Füge Ziele aus .env zuerst hinzu
        for target in normalized_env_targets:
            if self.combo_exportziel.findText(target) == -1:
                self.combo_exportziel.addItem(target)

        # Füge zuletzt gespeicherte Verzeichnisse hinzu
        recent_targets = self.loadRecentExportTargets()

        # Füge nur die Verzeichnisse hinzu, die nicht bereits in der .env-Datei enthalten sind
        #for target in recent_targets[:self.max_combobox_items - len(self.prefixed_targets)]:  # Kombinieren
        #    self.combo_exportziel.addItem(target)

        # Füge nur die Verzeichnisse hinzu, die nicht bereits in der .env-Datei enthalten sind
        for target in recent_targets:
            normalized_target = os.path.normpath(target).replace("\\", "/")
            if normalized_target not in normalized_env_targets and self.combo_exportziel.findText(normalized_target) == -1:
                self.combo_exportziel.addItem(normalized_target)

        # Füge den Eintrag für neues Ziel hinzu
        self.combo_exportziel.addItem(self.new_target_text)

        # Setze den Platzhalter-Text als aktiven Eintrag
        self.combo_exportziel.setCurrentIndex(0)


    def loadRecentExportTargets(self):
        """
        Lädt die letzten Exportziele aus der Datei.
        """
        if not os.path.exists(self.saved_export_file):
            return []

        with open(self.saved_export_file, "r", encoding="utf-8") as file:
            lines = file.readlines()
            return [line.strip() for line in lines if line.strip()]  # Leerzeilen entfernen


    def saveExportTarget(self, target):
        """
        Speichert ein neues Exportziel in der Historie, falls es nicht bereits existiert.
        """
        recent_targets = self.loadRecentExportTargets()

        # Verhindere Duplikate
        if target in recent_targets:
            recent_targets.remove(target)
        recent_targets.insert(0, target)  # Neues Ziel an den Anfang setzen

        # Historie auf max. Anzahl begrenzen
        recent_targets = recent_targets[:self.max_saved_targets]

        # Schreibe aktualisierte Historie zurück in die Datei
        with open(self.saved_export_file, "w", encoding="utf-8") as file:
            file.write("\n".join(recent_targets))


    def onExportTargetChanged(self, index):
        """
        Reaktion auf die Änderung in der ComboBox-Auswahl.
        """
        app_logger.trace(f"onExportTargetChanged wurde aufgerufen. Index: {index}, Typ: {type(index)}")

        if index < 0:
            return  # Verhindere Verarbeitung für ungültige Indizes

        target = self.combo_exportziel.currentText()
        app_logger.debug(f"Gewählter Eintrag: {target}")

        # Prüfe, ob Platzhalter ausgewählt wurde
        if target == self.placeholder_text:
            return  # Keine Aktion notwendig

        # Überprüfen, ob der Benutzer "Neues Exportziel wählen..." gewählt hat
        if target == self.new_target_text:
            self.handleNewTarget()
            return

        # Entferne Platzhalter, falls ein gültiges Ziel gewählt wird
        idx = self.combo_exportziel.findText(self.placeholder_text)
        if idx != -1:
            self.combo_exportziel.removeItem(idx)

        # Ziel ins Speicherprotokoll und Liste übernehmen
        self.saveExportTarget(target)


    def handleNewTarget(self):
        """
        Öffnet einen Dialog zum Auswählen eines neuen Exportziels.
        """
        # Temporäres Blockieren des Signals, um endlose Zyklen zu vermeiden
        self.combo_exportziel.blockSignals(True)

        try:
            # Debug-Ausgabe: Start des Dialogs
            app_logger.debug("Öffne QFileDialog zum Auswählen eines neuen Exportziels...")

            # QFileDialog mit korrektem Parent setzen
            dialog = QFileDialog(self.gui)  # Parent ist das Haupt-Widget (GUI)
            dialog.setFileMode(QFileDialog.Directory)  # Nur Verzeichnisse auswählen
            dialog.setOption(QFileDialog.ShowDirsOnly, True)  # Keine Dateien anzeigen
            dialog.setWindowTitle("Wählen Sie ein Exportziel aus")  # Dialogtitel setzen

            # Debug-Ausgabe: Dialog geöffnet?
            app_logger.debug("QFileDialog geöffnet.")

            # Zeige Dialog und prüfe, ob der Benutzer etwas ausgewählt hat
            if dialog.exec():
                selected_dir = dialog.selectedFiles()[0]  # Verzeichnisauswahl (immer 1 bei QFileDialog)
                app_logger.debug(f"Ausgewähltes Verzeichnis: {selected_dir}")  # Debug-Ausgabe

                if selected_dir:  # Nur, wenn ein Verzeichnis ausgewählt wurde
                    # Normalisiere den Pfad
                    normalized_dir = os.path.normpath(selected_dir).replace("\\", "/")
                    app_logger.debug(f"Ausgewähltes Verzeichnis (normalisiert): {normalized_dir}")

                    # Prüfen, ob das Verzeichnis bereits in der ComboBox ist
                    idx = self.combo_exportziel.findText(selected_dir)
                    if idx == -1:
                        # Neuer Eintrag zur combo_exportziel hinzufügen vor "Neues Exportziel wählen..."
                        self.combo_exportziel.insertItem(self.combo_exportziel.count() - 1, normalized_dir)
                        app_logger.debug(f"Neues Verzeichnis hinzugefügt: {normalized_dir}")  # Debug-Ausgabe

                    # Speicher den Eintrag in der Historie
                    self.saveExportTarget(normalized_dir)
                    app_logger.debug(f"Verzeichnis in Historie gespeichert: {normalized_dir}")  # Debug-Ausgabe

                    # Den neuen Eintrag in der ComboBox selektieren
                    self.combo_exportziel.setCurrentText(normalized_dir)

            else:
                app_logger.debug("Der QFileDialog wurde abgebrochen.")  # Debug-Ausgabe für Abbruch

        except Exception as e:
            app_logger.warning(f"Fehler im Dialog: {e}")  # Fehlerprotokollierung für Debugging

        finally:
            # Reaktiviere Signale
            self.combo_exportziel.blockSignals(False)

    def connectSignals(self):
        """
        Verbindet Ereignisse der combo_exportziel mit den entsprechenden Funktionen.
        """
        if self.combo_exportziel is None:
            app_logger.debug("combo_exportziel ist None. Überprüfen Sie die GUI-Referenz.")
        else:
            app_logger.debug("combo_exportziel ist korrekt initialisiert.")
            app_logger.debug("Versuche, Signal 'currentIndexChanged' mit 'onExportTargetChanged' zu verbinden...")
            self.combo_exportziel.currentIndexChanged.connect(self.onExportTargetChanged)
            app_logger.debug("Signal erfolgreich verbunden!")

    def connectSignals(self):
        """
        Verbindet Ereignisse der combo_exportziel mit den entsprechenden Funktionen.
        """
        app_logger.debug("Versuche, Signal 'currentIndexChanged' mit 'onExportTargetChanged' zu verbinden...")
        self.combo_exportziel.currentIndexChanged.connect(lambda index: self.onExportTargetChanged(index))

        app_logger.debug("Signal erfolgreich verbunden!")

        # Temporäre Signalprüfung mit Lambda für Debug-Zwecke
        self.combo_exportziel.currentIndexChanged.connect(
            lambda index: app_logger.debug(f"Signal ausgelöst! Indexänderung: {index}")
        )


def connect_gui_signals_exportziel_manager(gui):
    """
    Verbindet alle relevanten GUI-Signale mit den passenden Logik-Slots.
    """
    exportziel_manager = ExportzielManager(gui) # Erstelle eine Instanz von ExportzielManager
    exportziel_manager.connectSignals()  # Verbindet Signale