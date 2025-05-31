"""
Modul: exportziel_manager.py

Dieses Modul dient der Verwaltung von Exportzielen in einer GUI-Anwendung. Es kombiniert die
Pfad-Validierung, Signalverwaltung und Benutzerinteraktionslogik in einer zentralen Komponente.
Das Modul sorgt dafür, dass Benutzereingaben (z. B. Exportziele) validiert und verarbeitet werden
und ermöglicht eine nahtlose Integration in die GUI.

Hauptfunktionen:
-----------------
1. **Klasse `ExportzielManager`:**
   - Kernklasse für die Verwaltung und Verarbeitung von Exportzielen in der Anwendung.
   - Verantwortlich für die Verbindung zwischen GUI-Komponenten (z. B. ComboBox) und der notwendigen Logik.
   - Implementiert Signal-Slot-Mechanismen, um Benutzerinteraktionen zu verarbeiten.

2. **Pfad-Verarbeitungsmethoden:**
   - `ist_zulaessiger_pfad(pfad)`: Überprüft, ob ein angegebener Pfad entweder eine gültige lokale Pfadangabe
     oder eine WebDAV/HTTP(S)-URL ist.
   - `normalisiere_pfad(pfad)`: Wandelt Pfade in ein konsistentes und standardisiertes Format um.

3. **Signalverwaltung:**
   - Bindet Benutzeraktionen (z. B. Änderungen in einer ComboBox) an spezifische Ereignishandler.
   - Stelle sicher, dass die Benutzeroberfläche dynamisch auf Änderungen reagieren kann.

4. **Hilfsmethoden und Protokollierung:**
   - Protokollierung von Benutzeraktionen und Fehlern mithilfe des `app_logger`.
   - Unterstützende Funktionen für das Laden und Speichern von Exportzielen sowie für die Validierung von Eingaben.

Abhängigkeiten:
---------------
- **`os`**: Wird für die Arbeit mit Dateipfaden und deren Normalisierung verwendet.
- **`urllib.parse`**: Unterstützt die Validierung und Verarbeitung von WebDAV/HTTP(S)-URLs.
- Eine Logger-Instanz (`app_logger`), um Ereignisse und Fehler korrekt zu protokollieren.

Verwendungszweck:
-----------------
Das Modul wird für die Verwaltung und Bearbeitung von Exportzielen in einer GUI-Anwendung genutzt.
Es stellt sicher, dass Benutzereingaben effektiv verarbeitet, validiert und gespeichert werden.
Mit der Signal-Slot-Logik ermöglicht es eine reibungslose Integration in die Benutzeroberfläche
und die dynamische Reaktion auf Interaktionen.

Beispiel:
---------
Die folgende Code-Skizze zeigt die typische Verwendung des Moduls in einer GUI-Anwendung:

    from exportziel_manager import ExportzielManager

    # ExportzielManager-Instanz initialisieren und Signale verbinden
    exportziel_manager = ExportzielManager(gui)
    exportziel_manager.connectSignals()

    # Validierung und Normalisierung von Pfaden
    pfad = "C:/Benutzer/Dokumente/Export"
    if exportziel_manager.ist_zulaessiger_pfad(pfad):
        print("Pfad ist gültig.")
        normalisierter_pfad = exportziel_manager.normalisiere_pfad(pfad)
        print(f"Normalisierter Pfad: {normalisierter_pfad}")
    else:
        print("Ungültiger Pfad.")

Hinweise:
---------
- Das Modul verarbeitet sowohl lokale Pfade als auch WebDAV/HTTP(S)-URLs.
- Es wird empfohlen, bestehende Benutzerprofile oder Konfigurationsdateien zu nutzen,
  um häufig verwendete Exportziele zu speichern und wiederzuverwenden.
- Für komplexe Validierungen (z. B. Netzwerklaufwerke oder spezifische Berechtigungen)
  sollte die Methode `ist_zulaessiger_pfad` angepasst oder erweitert werden.
"""
import os
import logging
from urllib.parse import urlparse
from PySide6.QtWidgets import QFileDialog
from PySide6.QtCore import Qt
from dotenv import load_dotenv

from config import MAX_ENV_EXPORT_TARGETS, MAX_SAVED_EXPORT_TARGETS

# Erstellen eines Loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

# Globale Referenz für die Singleton-Instanz des ExportzielManagers
# Diese Referenz wird verwendet, um sicherzustellen, dass immer dieselbe Instanz verwendet wird.
exportziel_manager_instance = None 

# Schutzvariable, um zyklische Aufrufe und Rekursion bei der Initialisierung zu verhindern
# (z.B. während der Erstellung oder Nutzung des ExportzielManagers).
is_initializing = False 

class ExportzielManager:
    
    def __init__(self, gui):
        """
        Konstruktor der Klasse.

        Diese Methode initialisiert eine Instanz der Klasse und richtet alle erforderlichen Ressourcen und Verbindungen ein.
        Sie wird aufgerufen, wenn ein neues Objekt dieser Klasse erstellt wird.

        Parameter:
        ----------
        gui : object
            Die übergebene GUI-Instanz, die verwendet wird, um die grafischen Elemente (z. B. Widgets, Buttons, ComboBoxen etc.)
            zu verwalten und mit Logik zu verbinden.

        Ablauf:
        -------
        - Speichert die übergebene GUI-Referenz als Instanzvariable, um auf alle GUI-Elemente zugreifen zu können.
        - Konfiguriert und initialisiert die Benutzeroberfläche (z. B., ComboBox-Einträge oder andere GUI-Elemente).
        - Lädt globale oder Umgebungsdaten, die benötigt werden, um die Instanz korrekt zu konfigurieren.
        - Verbindet GUI-Signale (Benutzeraktionen wie Klicks oder Auswahländerungen) mit entsprechenden Logikmethoden.

        Hinweise:
        ---------
        - Dieser Konstruktor sollte nur einmal ausgeführt werden.
        - Alle initialisierten Instanzvariablen sollten umfassend kommentiert werden, um ihre Rolle im Programm zu erklären.
        """
        # Schreibe eine Protokollmeldung zum Start der Initialisierung der ComboBox (für Debugging-Zwecke).
        app_logger.debug(f"Start Initialisierung der ComboBox für die Auswahl von Exportzielen...")

        # Eine Referenz zur GUI-Instanz wird gespeichert, um Zugriff auf die zugehörigen GUI-Komponenten zu ermöglichen.
        self.gui = gui

        # Setzt eine Referenz auf die ComboBox zur Auswahl von Exportzielen aus der GUI.
        # Diese ComboBox wird verwendet, um Exportziele anzuzeigen und vom Nutzer auswählen zu lassen.
        self.combo_exportziel = self.gui.combo_exportziel

        # Definiert den Dateipfad, in dem die zuletzt genutzten Exportverzeichnisse gespeichert werden.
        self.saved_export_file = "export_history.txt"

        # Legt die maximale Anzahl von gespeicherten Exportzielen fest, die in der Historie aufbewahrt werden dürfen.
        self.max_saved_targets = MAX_SAVED_EXPORT_TARGETS

        # Definiert die maximale Anzahl von Exportzielen, die aus den `.env`-Dateien (Umgebungsvariablen) geladen werden können.
        self.max_env_targets = MAX_ENV_EXPORT_TARGETS

        # Platzhaltertext, der in der ComboBox angezeigt wird, wenn kein Exportziel ausgewählt wurde.
        # Dieser Text dient als hilfreiche Anleitung für den Benutzer.
        self.placeholder_text = "Bitte vor Start des Exports ein Zielverzeichnis wählen..."

        # Text, der in der ComboBox angezeigt wird, um dem Benutzer die Option zu geben, ein neues Exportziel hinzuzufügen.
        self.new_target_text = "Neues Exportziel wählen..."

        # Protokollmeldung zur Bestätigung, dass die Instanz des ExportzielManagers erfolgreich erstellt wurde.
        app_logger.debug("ExportzielManager wurde instanziiert.")
        

    def initExportTargets(self):
        """
        Initialisiert die Exportziele in der ComboBox.

        Der Ablauf umfasst:
        1. Laden der Ziele aus der `.env`-Datei.
        2. Normalisierung und Überprüfung der Pfade.
        3. Laden und Einfügen der zuletzt verwendeten Exportziele.
        4. Einfügen eines Trennstrichs (nicht auswählbar).
        5. Hinzufügen einer Option für ein neues Zielverzeichnis.
        6. Verbinden des Signals zur Auswahländerung mit einer Prüfungs-Methode.
        """
        # Leeren und Platzhalter setzen
        self.combo_exportziel.clear()
        self.combo_exportziel.addItem(self.placeholder_text)

        app_logger.debug(f"Jetzt ComboBox für die Auswahl von Exportzielen mit Einträgen befüllen...")

        # Schritt 1: Lade Ziele aus der `.env`-Datei
        load_dotenv()  # Lädt die Konfiguration aus `.env`
        prefixed_targets = []
        for i in range(1, self.max_env_targets + 1):
            # Hole den Wert der Umgebungsvariable "EXPORT_TARGET_i", wobei "i" die Nummer des Zielpfades ist
            target = os.getenv(f"EXPORT_TARGET_{i}")
            
            # Überprüfen, ob die Umgebungsvariable gesetzt ist und der Pfad gültig ist
            if target and self.ist_zulaessiger_pfad(target):  
                # Normalisiere den gültigen Pfad, um sicherzustellen, dass er konsistent und einheitlich formatiert ist
                normalisiert = self.normalisiere_pfad(target)
                
                # Füge den normalisierten und überprüften Pfad der Liste der Exportziele hinzu
                prefixed_targets.append(normalisiert)

        if prefixed_targets:
            # Wenn die Liste der vordefinierten Exportziele nicht leer ist, bearbeite sie weiter
            for target in sorted(prefixed_targets):  
                # Füge jedes Ziel aus der Liste, alphabetisch sortiert, als neuen Eintrag zur ComboBox hinzu
                self.combo_exportziel.addItem(target)

        app_logger.debug(f"Einträge aus vordefinierten Exportzielen: {prefixed_targets}")

        # Schritt 2: Lade zuletzt genutzte Exportziele
        recent_targets = self.loadRecentExportTargets()

        # Lade die Liste der zuletzt verwendeten Exportziele
        recent_targets = self.loadRecentExportTargets()

        # Prüfe, ob es Einträge in den zuletzt verwendeten Exportzielen gibt
        if recent_targets:
            # Füge einen visuellen Trennstrich in die ComboBox ein, um die vordefinierten Exportziele 
            # von den zuletzt verwendeten zu trennen
            self.combo_exportziel.addItem("-----------------")
            
            # Erhalte den Index des gerade hinzugefügten Trennstrichs und deaktiviere ihn,
            # sodass er nicht auswählbar ist
            index = self.combo_exportziel.count() - 1
            self.combo_exportziel.setItemData(index, Qt.NoItemFlags)
            
            # Iteriere über alle zuletzt verwendeten Exportziele
            for target in recent_targets:
                # Füge nur Ziele hinzu, die nicht schon in 'prefixed_targets' enthalten sind
                # und einen gültigen Pfad darstellen
                if target not in prefixed_targets and self.ist_zulaessiger_pfad(target):
                    # Normalisiere den gültigen Pfad für Konsistenz
                    normalisiert = self.normalisiere_pfad(target)
                    # Füge das normalisierte Ziel der ComboBox hinzu
                    self.combo_exportziel.addItem(normalisiert)

        app_logger.debug(f"Einträge aus historischen Exportzielen: {recent_targets}")

        # Schritt 3: Füge Option für Auswahl eines neuen Ziels hinzu
        self.combo_exportziel.addItem(self.new_target_text)

        # Schritt 4: Verbinde das Signal mit der Auswahlprüfung
        self.combo_exportziel.currentIndexChanged.connect(self.on_combobox_changed)

        # Platzhalter-Auswahl aktivieren
        self.combo_exportziel.setCurrentIndex(0)

        app_logger.debug(f"ComboBox für die Auswahl von Exportzielen ist mit {self.combo_exportziel.count()} Einträgen befüllt.")


    def on_combobox_changed(self, index):
        """
        Verarbeitet Änderungen in der Exportziel-ComboBox.

        Diese Methode wird aufgerufen, wenn der Benutzer eine Auswahl in der
        ComboBox (`combo_exportziel`) ändert. Sie verhindert ungültige
        Auswahlen, wie z. B. den Trennstrich, und aktualisiert den Zustand
        der zugehörigen Export-Buttons.

        Parameter:
        ----------
        index : int
        Der Index des aktuell ausgewählten Elements in der ComboBox.

        Ablauf:
        -------
        1. Der Text des aktuell ausgewählten Elements wird basierend auf dem
           übergebenen Index abgerufen.
        2. Wenn der ausgewählte Eintrag ein ungültiges Element wie der 
           Trennstrich ist, wird die Auswahl zurückgesetzt.
        3. Die Verfügbarkeit der Export-Buttons wird aktualisiert, um 
           sicherzustellen, dass sie nur bei gültigen Auswahlen aktiviert sind.
        4. Die getroffene Auswahl wird im Log hinterlegt, um die Benutzer-
           Interaktion zu verfolgen.
        """
        from gui_controller import update_export_buttons_state

        # Hole den aktuell ausgewählten Text basierend auf dem Index
        selected_item = self.combo_exportziel.itemText(index)

        # Überprüfe, ob der ausgewählte Eintrag der Trennstrich ist
        if selected_item == "-----------------":
            # Setze die Auswahl zurück auf den ersten Eintrag (Platzhalter oder Standard-Wert)
            self.combo_exportziel.setCurrentIndex(0)
            return  # Beende die Methode, da ein gültiger Eintrag nicht vorliegt

        # Export-Buttons initial deaktivieren
        update_export_buttons_state(self.gui)

        # Protokolliere die getroffene Auswahl für Debugging-Zwecke
        app_logger.debug(f"Auswahl: {selected_item}")


    def loadRecentExportTargets(self):
        """
        Liest die zuletzt verwendeten Exportziele aus einer Datei und gibt diese zurück.

        Diese Methode dient dazu, die Historie der zuletzt verwendeten Exportziele des Nutzers aus 
        einer lokal gespeicherten Datei (`self.saved_export_file`) auszulesen. Falls die Datei 
        nicht existiert, wird eine leere Liste zurückgegeben, um mögliche Fehler beim Lesen zu vermeiden.

        Rückgabewert:
        -------------
        list:
        - Eine Liste von Zeichenketten (Strings), die die zuletzt verwendeten 
          Exportziel-Verzeichnisse repräsentieren.
        - Falls die Datei fehlt oder keine gültigen Einträge enthält, wird eine leere 
          Liste zurückgegeben.

        Ablauf:
        -------
        1. Überprüfung, ob die Datei mit den gespeicherten Exportzielen existiert.
        2. Bei Nicht-Existenz der Datei Rückgabe einer leeren Liste.
        3. Datei mit UTF-8-Encoding öffnen und Zeilen auslesen.
        4. Leerzeichen an den Rändern jeder Zeile entfernen und leere Zeilen überspringen.
        5. Rückgabe einer Liste mit bereinigten Exportzielen.
        """
        # Überprüfe, ob die Datei existiert, in der die Exportziele gespeichert sind.
        # Wenn die Datei nicht existiert, wird eine leere Liste zurückgegeben, da keine Werte gelesen werden können.
        if not os.path.exists(self.saved_export_file):
            return []  # Rückgabe einer leeren Liste bei nicht existierender Datei.

        # Öffne die Datei mit UTF-8-Encoding, um ihre Inhalte im Lesemodus zu lesen.
        with open(self.saved_export_file, "r", encoding="utf-8") as file:
            # Lese alle Zeilen aus der Datei in die `lines`-Liste.
            lines = file.readlines()
            
            # Verarbeite die gelesenen Inhalte:
            # - Entferne führende und schließende Leerzeichen von jeder Zeile (strip()).
            # - Filtere alle Zeilen heraus, die komplett leer sind.
            return [line.strip() for line in lines if line.strip()]


    def saveExportTarget(self, target):
        """
        Fügt ein neues Exportziel zur Historie hinzu und speichert diese, 
        falls das Ziel nicht bereits vorhanden ist.

        Details:
        - Bestehende Exportziele werden so verarbeitet, dass Duplikate vermieden werden.
        - Das neue Ziel wird am Anfang der Historie eingefügt.
        - Die Historie wird auf eine maximale Anzahl an Einträgen (`max_saved_targets`) begrenzt.
        - Der aktuelle Zustand der Historie wird in einer Datei gespeichert.

        Parameter:
        - target (str): Das neue Exportziel, das gespeichert werden soll.
        """
        # Lade die bisherigen Exportziele aus der Datei
        recent_targets = self.loadRecentExportTargets()

        # Entferne das Ziel, falls es bereits existiert, um Duplikate zu vermeiden
        if target in recent_targets:
            recent_targets.remove(target)

        # Füge das neue Ziel am Anfang der Liste hinzu
        recent_targets.insert(0, target)

        # Kürze die Liste auf die maximale erlaubte Anzahl von Zielen
        recent_targets = recent_targets[:self.max_saved_targets]

        # Schreibe die aktualisierte Historie der Exportziele in die Datei
        with open(self.saved_export_file, "w", encoding="utf-8") as file:
            file.write("\n".join(recent_targets))


    def onExportTargetChanged(self, index):
        """
        Wird aufgerufen, wenn die Auswahl eines Exportziels in der ComboBox geändert wird.

        Zweck:
        ------
        Diese Methode reagiert auf die Änderung der Auswahl in der Exportziel-ComboBox.
        Sie prüft die Gültigkeit der Auswahl, führt je nach ausgewähltem Ziel spezifische Aktionen aus
        und aktualisiert konfigurationsrelevante Elemente.

        Ablauf:
        -------
        1. Protokolliert den Aufruf und den übermittelten Index für Debugging-Zwecke.
        2. Ignoriert ungültige Indizes, wie -1.
        3. Prüft auf speziell definierte Einträge (Platzhalter oder "Neues Ziel wählen").
        4. Entfernt den Platzhalter aus der Liste, falls ein reguläres Ziel gewählt wurde.

        Parameter:
        ----------
        index : int
        Der Index des neu ausgewählten Eintrags in der ComboBox.
        """
        # Schreibe eine Debug-Meldung mit Informationen zur Methode und dem Index (zur Nachverfolgung von Benutzeraktionen).
        app_logger.debug(f"onExportTargetChanged wurde aufgerufen. Index: {index}, Typ: {type(index)}")

        # Ignoriere den Aufruf, wenn der Index ungültig ist (z. B. bei keiner Auswahl oder Rückstellung).
        if index < 0:
            return

        # Hole den aktuell ausgewählten Text aus der ComboBox, um zu prüfen, welche Aktion ausgelöst werden soll.
        target = self.combo_exportziel.currentText()
        app_logger.debug(f"Der Eintrag '{target}' wurde ausgewählt.")

        # Überspringe die Verarbeitung, wenn der Platzhalter-Eintrag ausgewählt wurde.
        if target == self.placeholder_text:
            app_logger.debug(f"Die Auswahl des Platzhalters '{target}' wird ignoriert. Kein Exportziel wurde festgelegt.")
            return

        # Behandlung eines speziellen Textes (z. B. "Neues Exportziel wählen..."):
        # Öffnet einen Dialog oder führt eine Aktion aus, um ein neues Exportziel hinzuzufügen.
        if target == self.new_target_text:
            app_logger.debug("Dialog zur Auswahl eines neuen Exportziels wird gestartet.")
            self.handleNewTarget()

        # Prüfen, ob der Platzhalter nach wie vor in der ComboBox existiert, und ihn entfernen, da ein gültiges Ziel gewählt wurde.
        idx = self.combo_exportziel.findText(self.placeholder_text)
        if idx != -1:
            self.combo_exportziel.removeItem(idx)

        # Protokolliert das neu ausgewählte und aktive Exportziel.
        app_logger.debug(f"Neues gewähltes Exportziel: {self.combo_exportziel.currentText()}")


    def handleNewTarget(self):
        """
        Öffnet einen Dialog zum Auswählen eines neuen Exportziels und aktualisiert die Export-Historie.

        Zweck:
        ------
        Diese Methode ermöglicht dem Benutzer, ein neues Exportziel-Verzeichnis auszuwählen.
        Das gewählte Ziel wird in der ComboBox angezeigt, in der Historie gespeichert
        und bei einer gültigen Auswahl als aktuelles Exportziel eingestellt.
        
        Ablauf:
        -------
        1. Blockiert vorübergehend die Signals der ComboBox, um Rekursionen oder unnötige Signalauslöser zu verhindern.
        2. Öffnet den QFileDialog, damit der Benutzer ein Verzeichnis auswählen kann.
        3. Überprüft, ob ein gültiges Verzeichnis ausgewählt wurde:
            - Normalisiert den Pfad zur Standardform.
            - Fügt das Verzeichnis zur ComboBox hinzu, falls es noch nicht vorhanden ist.
            - Speichert das Verzeichnis in der Exportziel-Historie.
            - Wählt das neue Verzeichnis als aktives Ziel in der ComboBox aus.
        4. Bei Abbruch oder Fehler werden relevante Protokolle geschrieben, und die ComboBox-Signale
           werden wieder aktiviert.

        Besonderheiten:
        ---------------
        - Es wird geprüft, ob das ausgewählte Verzeichnis bereits in der ComboBox existiert.
        - Das Hinzufügen eines neuen Ziels erfolgt vor Einträgen wie "Neues Exportziel wählen...", um
          die Benutzerfreundlichkeit zu gewährleisten.
        - Die Methode stellt sicher, dass keine endlosen Ereigniszyklen durch Signalwirkungen ausgelöst werden.

        Behandlung im Fehlerfall:
        -------------------------
        Alle potenziellen Fehler bei der Auswahl eines Verzeichnisses oder beim Aktualisieren
        der GUI werden abgefangen und protokolliert.

        """
        # Blockiert vorübergehend Signals der ComboBox, um Rekursionen oder unerwünschte Ereignisse zu verhindern
        self.combo_exportziel.blockSignals(True)

        try:
            # Debug-Meldungen zur Überwachung des Prozesses
            app_logger.debug("Öffne QFileDialog zum Auswählen eines neuen Exportziels...")

            # Initiierung des QFileDialogs mit spezifischen Optionen
            dialog = QFileDialog(self.gui)  # Setzt das GUI-Hauptfenster als übergeordnetes Element
            dialog.setFileMode(QFileDialog.Directory)  # Begrenzung auf Verzeichnisauswahl
            dialog.setOption(QFileDialog.ShowDirsOnly, True)  # Nur Verzeichnisse anzeigen
            dialog.setWindowTitle("Wählen Sie ein Exportziel aus")  # Dialogtitel

            # Prüft, ob der Dialog erfolgreich ausgeführt wurde (Benutzer hat etwas ausgewählt)
            if dialog.exec():
                selected_dir = dialog.selectedFiles()[0]  # Erstes (und einziges) ausgewähltes Verzeichnis
                app_logger.debug(f"Ausgewähltes Verzeichnis: {selected_dir}")  # Debug-Ausgabe

                # Prüfen, ob das Verzeichnis gültig ist
                if selected_dir:
                    # Normalisiere den Pfad (z. B. ersetze Backslashes durch Forward Slashes für Plattform-Konsistenz)
                    normalized_dir = os.path.normpath(selected_dir).replace("\\", "/")
                    app_logger.debug(f"Ausgewähltes Verzeichnis (normalisiert): {normalized_dir}")

                    # Überprüfen, ob das Verzeichnis bereits in der ComboBox existiert
                    idx = self.combo_exportziel.findText(selected_dir)
                    if idx == -1:
                        # Füge das Verzeichnis zur ComboBox hinzu, vor dem Eintrag "Neues Exportziel wählen..."
                        self.combo_exportziel.insertItem(self.combo_exportziel.count() - 1, normalized_dir)
                        app_logger.debug(f"Neues Verzeichnis hinzugefügt: {normalized_dir}")

                # Das Verzeichnis wird in der Historie persistent gespeichert
                self.saveExportTarget(normalized_dir)
                app_logger.debug(f"Verzeichnis in Historie gespeichert: {normalized_dir}")

                # Setze das aktuelle Textfeld der ComboBox auf das neue Verzeichnis
                self.combo_exportziel.setCurrentText(normalized_dir)

            else:
                # Falls der Benutzer den Dialog abbricht
                app_logger.debug("Der QFileDialog wurde abgebrochen.")

        except Exception as e:
            # Fehlerbehandlung und Protokollierung
            app_logger.warning(f"Fehler beim Öffnen des Dialogs oder Verarbeiten des Verzeichnisses: {e}")

        finally:
            # Reaktivieren der ComboBox-Signale nach Abschluss aller Aktionen
            self.combo_exportziel.blockSignals(False)


    def connectSignals(self):
        """
        Verbindet die Signale der ComboBox `combo_exportziel` mit den zugehörigen Funktionen, 
        um Benutzerinteraktionen zu verarbeiten.

        Zweck:
        ------
        Diese Methode stellt sicher, dass Änderungen an der Auswahl in der ComboBox `combo_exportziel` 
        erkannt und verarbeitet werden, indem sie das Signal `currentIndexChanged` mit der entsprechenden 
        Methode `onExportTargetChanged` verbindet.

        Ablauf:
        -------
        1. Überprüft, ob die ComboBox `combo_exportziel` korrekt initialisiert wurde:
        - Falls die ComboBox nicht existiert (None), wird eine Debug-Warnung protokolliert.
        - Falls die ComboBox korrekt vorhanden ist, wird mit der Signalverbindung fortgefahren.
        2. Verbindet das Signal `currentIndexChanged` der ComboBox mit der Methode `onExportTargetChanged`,
        um Änderungen der Auswahl zu behandeln.
        3. Protokolliert den Status und die Ergebnisse der Verbindung zur Unterstützung des Debuggings.

        Besonderheiten:
        ---------------
        - Verbindet das Signal mit einer Lambda-Funktion, um den Index als Argument an die Methode `onExportTargetChanged` zu übergeben.
        - Unterstützt Debugging durch detaillierte Log-Ausgaben in allen Schritten.
        - Stellt sicher, dass die Signalverbindung nur bei einer korrekt initialisierten ComboBox erfolgt.

        Fehlerbehandlung:
        -----------------
        Es gibt keine direkten Fehlerbehandlungen innerhalb dieser Methode. Falls die ComboBox nicht 
        existiert, wird eine Debug-Warnung protokolliert, und die Signalverbindung wird nicht durchgeführt.

        """
        # Prüfen, ob die Referenz auf die ComboBox existiert
        if self.combo_exportziel is None:
            # Debug-Log: ComboBox ist nicht initialisiert
            app_logger.debug("combo_exportziel ist None. Überprüfen Sie die GUI-Referenz.")
        else:
            # Debug-Log: ComboBox wurde korrekt gefunden
            app_logger.debug("combo_exportziel ist korrekt initialisiert.")

            # Protokolliert den Beginn der Signalverbindung
            app_logger.debug("Versuche, Signal 'currentIndexChanged' mit 'onExportTargetChanged' zu verbinden...")

            # Signal `currentIndexChanged` mit der Methode `onExportTargetChanged` verbinden
            self.combo_exportziel.currentIndexChanged.connect(lambda index: self.onExportTargetChanged(index))

            # Debug-Log nach erfolgreicher Verbindungsherstellung
            app_logger.debug("Signal erfolgreich verbunden!")


    def ist_zulaessiger_pfad(self, pfad):
        """
        Überprüft, ob der angegebene Pfad ein gültiger lokaler Pfad oder eine gültige WebDAV/HTTP(S)-URL ist.

        Zweck:
        ------
        Diese Methode dient der Validierung von Pfaden, die entweder als WebDAV-/HTTP(S)-URLs oder 
        als absolute lokale Pfade angegeben werden können. Sie wird verwendet, um sicherzustellen, 
        dass nur zulässige Pfadformate akzeptiert werden.

        Ablauf:
        -------
        1. Zerlegt den angegebenen Pfad in seine Bestandteile (Parsing).
        2. Überprüft, ob das Schema des Pfads entweder "http" oder "https" ist (WebDAV/HTTP(S)-URL).
        3. Falls der Pfad keine URL ist, validiert er, ob es sich um einen absoluten lokalen Dateipfad handelt.
        4. Gibt `True` zurück, wenn einer der beiden Bedingungen erfüllt ist. 
           Andernfalls wird `False` zurückgegeben und der Pfad als ungültig eingestuft.

        Besonderheiten:
        ---------------
        - URLs werden durch die Prüfung des Schemas (`http`, `https`) validiert.
        - Lokale Pfade müssen absolut sein, um als gültig zu gelten.
        - Unterstützt keine relativen Pfade oder andere URL-Typen wie FTP oder File-URLs.

        Parameter:
        ----------
        pfad (str): Der Pfad oder die URL, die geprüft werden soll.

        Rückgabewert:
        -------------
        bool: 
        - `True`: Wenn der Pfad eine zulässige WebDAV/HTTP(S)-URL oder ein gültiger 
          absoluter lokaler Pfad ist.
        - `False`: Wenn der Pfad weder eine gültige URL noch ein gültiger lokaler Pfad ist.

        Beispiel:
        ---------
        - `ist_zulaessiger_pfad("https://example.com")` → `True`
        - `ist_zulaessiger_pfad("C:/Users/User/Documents")` → `True`
        - `ist_zulaessiger_pfad("invalid/path")` → `False`
        """
        # Zerlege den angegebenen Pfad in seine Bestandteile (Parsing)
        parsed = urlparse(pfad)  # Analysiert die Struktur des Pfads in Komponenten

        # Überprüfen, ob die Schema-Komponente eine gültige WebDAV/HTTP(S)-URL (http/https) darstellt
        if parsed.scheme in ("http", "https"):  
            # Gültige HTTP(S)-URL erkannt
            return True  

        # Überprüfen, ob der Pfad ein absoluter lokaler Pfad ist
        if os.path.isabs(pfad):  
            # Absoluter Pfad ist gültig
            return True  

        # Weder eine gültige URL noch ein absoluter lokaler Pfad → Rückgabe `False`
        return False


    def normalisiere_pfad(self, pfad):
        """
        Normalisiert den angegebenen Pfad abhängig von dessen Typ.

        - WebDAV/HTTP(S)-URLs bleiben unverändert, da sie direkt verwendet werden können.
        - Lokale Dateipfade werden normalisiert:
          * Entfernt überflüssige Separatoren und relative Elemente (z. B. "..", ".").
          * Ersetzt Backslashes (`\`) durch Slashes (`/`), um ein konsistentes Format zu gewährleisten.

        Ablauf:
        - Überprüft zunächst, ob der Pfad ein gültiges Format besitzt (lokaler Pfad oder WebDAV/HTTP(S)-URL).
          * Falls der Pfad ungültig ist, wird ein `ValueError` ausgelöst.
        - Gibt den Pfad unverändert zurück, wenn es sich um eine WebDAV/HTTP(S)-URL handelt.
        - Gibt den normalisierten lokalen Dateipfad im konsistenten Format zurück.

        Parameter:
        - pfad (str): Der zu normalisierende Pfad.

        Rückgabewert:
        - str: Der normalisierte Pfad.

        Ausnahme:
        - ValueError: Wird ausgelöst, wenn der angegebene Pfad kein gültiges Format hat.
        """
        # Überprüfen, ob der Pfad ein gültiger lokaler Pfad oder eine HTTP(S)-URL ist
        if not self.ist_zulaessiger_pfad(pfad):
            raise ValueError(f"Ungültiges Pfadformat: {pfad}")  # Fehler auslösen, falls der Pfad ungültig ist

        # Wenn es sich um eine WebDAV/HTTP(S)-URL handelt, den Pfad unverändert zurückgeben
        if pfad.startswith(("http://", "https://")):
            return pfad  # URL bleibt unverändert

        # Lokalen Dateipfad normalisieren und konsistente Ausgabe mit Slashes sicherstellen
        return os.path.normpath(pfad).replace("\\", "/")


def connect_gui_signals_exportziel_manager(gui):
    """
    Verbindet die relevanten GUI-Signale mit der Logik des Exportziel-Managers, 
    um Interaktionen zwischen der Benutzeroberfläche und dem Exportziel-Manager zu ermöglichen.

    Zweck:
    ------
    Diese Methode erstellt eine Instanz des `ExportzielManager`, verbindet die notwendigen 
    Signale der GUI mit den dazugehörigen Funktionen im Exportziel-Manager und stellt sicher, 
    dass die Instanz später in der GUI genutzt werden kann.

    Ablauf:
    -------
    1. Initialisiert eine `ExportzielManager`-Instanz mit der übergebenen GUI-Komponente `gui`.
    2. Ruft die Methode `connectSignals` des `ExportzielManager` auf, um relevante GUI-Signale 
       mit Funktionen des Managers zu verknüpfen.
    3. Weist die `ExportzielManager`-Instanz der GUI-Instanz (`gui.exportziel_manager`) zu, 
       damit diese in anderen Bereichen der GUI verfügbar ist.
    4. Gibt die Instanz des `ExportzielManager` zurück, falls sie für weitere Zwecke benötigt wird.

    Besonderheiten:
    ---------------
    - Diese Methode ist der zentrale Punkt für die Initialisierung und Signalverbindung des Exportziel-Managers.
    - Die Rückgabe des `ExportzielManager` ermöglicht eine flexible Weiterverarbeitung, falls nötig.
    - Setzt voraus, dass die GUI bereits korrekt initialisiert ist und die relevanten Widgets (z. B. ComboBox) vorhanden sind.

    Parameter:
    ----------
    gui: Die Hauptkomponente der GUI (z. B. Hauptfenster oder ein Widget), die mit dem Exportziel-Manager 
         interagiert.

    Rückgabewert:
    -------------
    ExportzielManager: Die neu erstellte und verbundene Instanz des Exportziel-Managers.

    Beispiel:
    ---------
    ```
    # Verbindet die GUI-Signale mit dem Exportziel-Manager
    exportziel_manager = connect_gui_signals_exportziel_manager(window)

    # Prüfen, ob die Instanz erfolgreich erstellt wurde
    if exportziel_manager:
        print("Exportziel-Manager erfolgreich konfiguriert.")
    ```
    """
    # Instanz des ExportzielManagers erstellen und Signale verbinden
    exportziel_manager = ExportzielManager(gui)  # Erstellt eine neue Instanz des `ExportzielManager`
    exportziel_manager.connectSignals()  # Verknüpft relevante Signale aus der GUI mit den zugehörigen Slots

    # Speichert die `ExportzielManager`-Instanz in der GUI-Instanz, damit sie wiederverwendet werden kann
    gui.exportziel_manager = exportziel_manager

    # Gibt die Instanz des ExportzielManagers zurück, falls sie später benötigt wird
    return exportziel_manager


def get_exportziel_manager(gui=None):
    """
    Gibt die Singleton-Instanz des `ExportzielManager` zurück.
    Wenn die Instanz noch nicht existiert, wird sie mit der übergebenen `gui`-Referenz neu erstellt.

    Zweck:
    ------
    Diese Funktion stellt sicher, dass der `ExportzielManager` als Singleton fungiert, indem 
    immer die gleiche Instanz zurückgegeben wird. Falls die Instanz noch nicht erstellt wurde, 
    wird sie mit der GUI-Referenz initialisiert.

    Ablauf:
    -------
    1. Prüft, ob die Singleton-Inst
    2. Ruft die Methode `connectSignals` des `ExportzielManager` auf, um relevante GUI-Signale
          mit Funktionen des Managers zu verknüpfen.
       3. Weist die `ExportzielManager`-Instanz der GUI-Instanz (`gui.exportziel_manager`) zu,
          damit diese in anderen Bereichen der GUI verfügbar ist.
       4. Gibt die Instanz des `ExportzielManager` zurück, falls sie für weitere Zwecke benötigt wird.

       Besonderheiten:
       ---------------
       - Diese Methode ist der zentrale Punkt für die Initialisierung und Signalverbindung des Exportziel-Managers.
       - Die Rückgabe des `ExportzielManager` ermöglicht eine flexible Weiterverarbeitung, falls nötig.
       - Setzt voraus, dass die GUI bereits korrekt initialisiert ist und die relevanten Widgets (z. B. ComboBox) vorhanden sind.

       Parameter:
       ----------
       gui: Die Hauptkomponente der GUI (z. B. Hauptfenster oder ein Widget), die mit dem Exportziel-Manager
            interagiert.

       Rückgabewert:
       -------------
       ExportzielManager: Die neu erstellte und verbundene Instanz des Exportziel-Managers.

       Beispiel:
       ---------
       ```
       # Verbindet die GUI-Signale mit dem Exportziel-Manager
       exportziel_manager = connect_gui_signals_exportziel_manager(window)

       # Prüfen, ob die Instanz erfolgreich erstellt wurde
       if exportziel_manager:
           print("Exportziel-Manager erfolgreich konfiguriert.")
       ```
       """
    # Instanz des ExportzielManagers erstellen und Signale verbinden
    exportziel_manager = ExportzielManager(gui)  # Erstellt eine neue Instanz des `ExportzielManager`
    exportziel_manager.connectSignals()  # Verknüpft relevante Signale aus der GUI mit den zugehörigen Slots

    # Speichert die `ExportzielManager`-Instanz in der GUI-Instanz, damit sie wiederverwendet werden kann
    gui.exportziel_manager = exportziel_manager

    # Gibt die Instanz des ExportzielManagers zurück, falls sie später benötigt wird
    return exportziel_manager
