# -*- coding: utf-8 -*-
"""
excel_handling.py

Dieses Modul enthält Funktionen zur Erstellung und Speicherung von Excel-Listen aus MSG-Dateien.
Es ermöglicht das Sammeln von Metadaten über MSG-Dateien in einem strukturierten Format, das leicht
in Excel exportiert werden kann. Die Hauptfunktionen umfassen das Erstellen einer Excel-Liste aus
einer Liste von MSG-Dateien sowie das Speichern dieser Liste in einer Excel-Datei.

Funktionen:
- create_excel_list(msg_files): Erstellt eine Excel-Liste aus den gefundenen MSG-Dateien.
- save_excel_file(excel_list, output_file): Speichert die Excel-Liste in einer angegebenen Datei.
"""

import os
import pandas as pd
import logging

# Erstellen eines Loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

def create_excel_list(msg_files):
    """
    Erstellt eine Excel-Liste aus den gefundenen MSG-Dateien.

    Diese Funktion erstellt ein DataFrame mit Informationen über jede gefundene MSG-Datei,
    einschließlich einer fortlaufenden Nummer, dem Dateinamen, dem Pfadnamen und der Länge des Pfades.

    :param msg_files: Eine Liste der gefundenen MSG-Dateien (vollständige Pfade).
    :return: Ein DataFrame, das die Informationen über die MSG-Dateien enthält.
    """
    # Initialisiere ein leeres DataFrame mit den entsprechenden Spalten
    excel_list = pd.DataFrame(columns=["Nummer", "Dateiname", "Pfadname", "Pfadlänge"])

    for i, file in enumerate(msg_files):
        filename = os.path.basename(file)  # Extrahiere den Dateinamen
        path = os.path.dirname(file)        # Extrahiere den Pfad
        entry = {
            "Nummer": i + 1,                # Fortlaufende Nummer
            "Dateiname": filename,           # Name der Datei
            "Pfadname": path,                # Verzeichnis der Datei
            "Pfadlänge": len(file)           # Länge des vollständigen Pfades
        }
        # Füge den neuen Eintrag zum DataFrame hinzu
        excel_list = pd.concat([excel_list, pd.DataFrame([entry])], ignore_index=True)
    return excel_list

def save_excel_file(excel_list, output_file):
    """
    Speichert die Excel-Liste in einer angegebenen Datei.

    Diese Funktion speichert das übergebene DataFrame als Excel-Datei an dem angegebenen Speicherort.

    :param excel_list: Das DataFrame, das gespeichert werden soll.
    :param output_file: Der Pfad zur Ausgabedatei, in der die Excel-Liste gespeichert wird.
    """
    # Speichere das DataFrame als Excel-Datei
    excel_list.to_excel(output_file, index=False)
    print(f"Excel-Liste erfolgreich gespeichert unter: {output_file}")