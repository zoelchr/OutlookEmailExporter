# -*- coding: utf-8 -*-
"""
file_handling.py

Dieses Modul enthält Funktionen zur Handhabung von Dateien und Verzeichnissen.
Es bietet Routinen zum Testen des Zugriffs auf Dateien, Umbenennen von Dateien,
Löschen von Dateien und Verzeichnissen sowie zum Kopieren von Inhalten zwischen Verzeichnissen.
Zusätzlich ermöglicht es das Setzen von Erstellungs- und Änderungsdaten für Dateien.

Funktionen:
- copy_directory_contents(source_directory_path, target_directory_path): Kopiert den gesamten Inhalt eines Quellverzeichnisses in ein Zielverzeichnis.
- delete_directory_contents(directory_path): Löscht den gesamten Inhalt eines angegebenen Verzeichnisses.
- delete_file(file_path: str, retries=1, delay_ms=200) -> FileOperationResult: Löscht eine Datei und gibt den Status zurück.
- format_datetime_stamp(datetime_stamp, format_string): Formatiert einen Zeitstempel in das angegebene Format.
- rename_file(current_name: str, new_name: str, retries=1, delay_ms=200) -> FileOperationResult: Benennt eine Datei um und gibt den Status zurück.
- sanitize_filename(filename: str) -> str: Ersetzt ungültige Zeichen durch Unterstriche.
- set_file_creation_date(file_path: str, new_creation_date: str) -> FileOperationResult: Setzt das Erstelldatum einer Datei auf einen vorgegebenen Wert und gibt den Status zurück.
- set_file_modification_date(file_path: str, new_date: str) -> FileOperationResult: Setzt das Änderungsdatum einer Datei auf einen vorgegebenen Wert und gibt den Status zurück.
- test_file_access(file_path: str) -> list[FileAccessStatus]: Überprüft den Datei-Zugriffsstatus und gibt eine Liste von FileAccessStatus-Enums zurück.

Verwendung:
Importieren Sie dieses Modul in Ihr Skript, um die oben genannten Funktionen zur Datei- und Verzeichnisverwaltung zu nutzen.
"""

import os
import shutil
import time
import pywintypes
import win32file
import win32con
import shutil
import re
import datetime
import logging
from enum import Enum

# Erstellen eines app_loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

class FileAccessStatus(Enum):
    READABLE = "File is readable"
    WRITABLE = "File is writable"
    EXECUTABLE = "File is executable"
    NOT_FOUND = "File not found"
    NO_PERMISSION = "No permission to access"
    LOCKED = "File is locked by another process"
    UNKNOWN_ERROR = "Unknown error"
    UNKNOWN = "Status is unkown"

class FileOperationResult(Enum):
    SUCCESS = "Success"
    FILE_NOT_FOUND = "File not found"
    DESTINATION_EXISTS = "Destination file already exists"
    TIMESTAMP_MATCH = "Timestamp match"
    PERMISSION_DENIED = "Permission denied"
    INVALID_FILENAME = "Invalid filename"
    INVALID_FILENAME1 = "Source is a file and destination is a directory"
    INVALID_FILENAME2 = "Part of the path is not a directory"
    UNKNOWN_ERROR = "Unknown error"
    VALUE_ERROR = "Value error"
    ERROR = "Error"

class FileHandle:
    """
    Kontextmanager für den Zugriff auf eine Datei mit Windows-API.

    Diese Klasse ermöglicht das sichere Öffnen und Schließen von Dateien unter Verwendung der Windows-API.
    Sie implementiert die Methoden __enter__ und __exit__, um sicherzustellen, dass die Datei
    ordnungsgemäß geöffnet und geschlossen wird, wenn sie in einem with-Block verwendet wird.

    Attribute:
    file_path (str): Der Pfad zur Datei, die geöffnet werden soll.
    handle (HANDLE): Der Handle für die geöffnete Datei, der von der Windows-API zurückgegeben wird.
    """

    def __init__(self, file_path: str):
        """
        Initialisiert ein neues FileHandle-Objekt.

        Parameter:
        file_path (str): Der Pfad zur Datei, die geöffnet werden soll.
        """
        self.file_path = file_path
        self.handle = None  # Handle wird initial auf None gesetzt

    def __enter__(self):
        """
        Öffnet die Datei und gibt den Handle zurück.

        Rückgabewert:
        HANDLE: Der Handle für die geöffnete Datei.
        """
        self.handle = win32file.CreateFile(
            self.file_path,
            win32con.GENERIC_WRITE,
            0,
            None,
            win32con.OPEN_EXISTING,
            win32con.FILE_ATTRIBUTE_NORMAL,
            None
        )
        return self.handle  # Gibt den Handle zurück, um ihn im with-Block verwenden zu können

    def __exit__(self, exc_type, exc_val, exc_tb):
        """
        Schließt den Datei-Handle, wenn der with-Block verlassen wird.

        Parameter:
        exc_type: Der Typ der Ausnahme, falls eine Ausnahme aufgetreten ist.
        exc_val: Der Wert der Ausnahme, falls eine Ausnahme aufgetreten ist.
        exc_tb: Der Traceback der Ausnahme, falls eine Ausnahme aufgetreten ist.
        """
        if self.handle is not None:
            win32file.CloseHandle(self.handle)  # Schließt den Handle, um Ressourcen freizugeben


def test_file_access(file_path: str) -> list[FileAccessStatus]:
    """Überprüft den Datei-Zugriffsstatus und gibt eine Liste von FileAccessStatus-Enums zurück.

    Beispiel:
    status = check_file_access(file_path)
    print(f"Status: {[s.value for s in status]}")
    """
    access_status = []

    try:
        # Prüfen, ob die Datei existiert
        if not os.path.exists(file_path):
            app_logger.warning(f"Datei nicht gefunden: {file_path}")
            return [FileAccessStatus.NOT_FOUND]

        # Prüfen, ob die Datei gesperrt ist (Windows-typisch)
        try:
            with open(file_path, "a"):  # Testweise öffnen zum Schreiben
                pass
        except PermissionError:
            app_logger.warning(f"Datei ist gesperrt oder nicht schreibbar: {file_path}")
            access_status.append(FileAccessStatus.LOCKED)

        # Prüfen, welche Berechtigungen vorliegen
        if os.access(file_path, os.R_OK):
            access_status.append(FileAccessStatus.READABLE)

        if os.access(file_path, os.W_OK):
            access_status.append(FileAccessStatus.WRITABLE)

        if os.access(file_path, os.X_OK):
            access_status.append(FileAccessStatus.EXECUTABLE)

        # Falls keine der Berechtigungen vorhanden ist
        if not access_status:
            app_logger.warning(f"Kein Zugriff auf Datei: {file_path}")
            access_status.append(FileAccessStatus.NO_PERMISSION)

        return access_status

    except FileNotFoundError:
        app_logger.error(f"Datei wurde nicht gefunden: {file_path}")
        return [FileAccessStatus.NOT_FOUND]

    except Exception as e:
        app_logger.exception(f"Unerwarteter Fehler beim Zugriff auf {file_path}: {e}")
        return [FileAccessStatus.UNKNOWN_ERROR]


def rename_file(current_name: str, new_name: str, retries=3, delay_ms=200, max_console_output=False) -> FileOperationResult:
    """
    Benennt eine Datei um und prüft die erfolgreiche Umbenennung.
    Neue verbesserte Version

    Parameter:
    current_name (str): Der aktuelle Dateiname.
    new_name (str): Der neue Dateiname.
    retries (int): Anzahl der Wiederholungen bei Misserfolg (Standard: 3).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 200).

    Rückgabewert:
    RenameResult: Enum-Wert, der den Erfolg oder Fehler beschreibt.
    """
    attempt = 0
    rename_file_result = None  # Variable für den Rückgabewert

    try:
        if max_console_output: print(f"Prüfe: {os.path.exists(current_name)}")
    except FileNotFoundError as e:
        if max_console_output: print("Datei nicht gefunden:", current_name)
        if max_console_output: print(e)
    except Exception as e:
        if max_console_output: print("Allgemeiner Fehler:", e)

    while attempt < retries:
        try:
            app_logger.debug(f"Aktueller Dateiname: {current_name}")  # Debugging-Ausgabe: Log-File
            app_logger.debug(f"Neuer Dateiname: {new_name}")  # Debugging-Ausgabe: Log-File
            #os.rename(current_name, new_name)
            shutil.move(current_name, new_name)
            rename_file_result = FileOperationResult.SUCCESS  # Erfolgreiche Umbenennung
            break # Erfolgreich umbenannt (ACHTUNG: Das muss auch noch im Programm msg_file_renamer korrigiert werden

        except FileNotFoundError:
            if max_console_output: print(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Console
            app_logger.error(f"Datei nicht gefunden: {current_name}")  # Debugging-Ausgabe: Log-File
            rename_file_result = FileOperationResult.FILE_NOT_FOUND  # Datei nicht gefunden
        except PermissionError:
            if max_console_output: print(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Console
            app_logger.error(f"Berechtigungsfehler bei Zugriff auf Datei: {current_name}")  # Debugging-Ausgabe: Log-File
            rename_file_result = FileOperationResult.PERMISSION_DENIED  # Berechtigungsfehler
        except FileExistsError:
            if max_console_output: print(f"Zieldatei existiert bereits: {new_name}")  # Debugging-Ausgabe: Console
            app_logger.error(f"Zieldatei existiert bereits: {new_name}")  # Debugging-Ausgabe: Log-File
            return FileOperationResult.DESTINATION_EXISTS
        except IsADirectoryError:
            if max_console_output: print(f"Kann nicht umbenennen, da die Quelle eine Datei und das Ziel ein Verzeichnis ist.")  # Debugging-Ausgabe: Console
            app_logger.error(f"Kann nicht umbenennen, da die Quelle eine Datei und das Ziel ein Verzeichnis ist.")  # Debugging-Ausgabe: Log-File
            return FileOperationResult.INVALID_FILENAME1
        except NotADirectoryError:
            if max_console_output: print(f"Ein Teil des Pfades ist kein Verzeichnis: {current_name} oder {new_name}")  # Debugging-Ausgabe: Console
            app_logger.error(f"Ein Teil des Pfades ist kein Verzeichnis: {current_name} oder {new_name}")  # Debugging-Ausgabe: Log-File
            return FileOperationResult.INVALID_FILENAME2
        except Exception as e:
            rename_file_result = FileOperationResult.UNKNOWN_ERROR  # Unbekannter Fehler
            app_logger.error(f"Fehler bei Umbenennung: {str(e)}")  # Debugging-Ausgabe: Log-File

        attempt += 1
        time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    return rename_file_result


def delete_file(file_path: str, retries=1, delay_ms=200) -> FileOperationResult:
    """
    Löscht eine Datei und prüft die erfolgreiche Löschung.

    Diese Funktion versucht, die angegebene Datei zu löschen.
    Bei einem Misserfolg wird die Löschung nach einer konfigurierbaren Anzahl von Versuchen wiederholt.
    Der Benutzer kann auch eine Verzögerung zwischen den Versuchen angeben.

    Parameter:
    file_path (str): Der Pfad zur Datei, die gelöscht werden soll.
    retries (int): Die Anzahl der Wiederholungen bei Misserfolg (Standard: 1).
    delay_ms (int): Millisekunden zwischen den Wiederholungen (Standard: 200).

    Rückgabewert:
    FileOperationResult: Ein Enum-Wert, der den Erfolg oder Fehler beschreibt.

    Beispiel:
        result = delete_file2('example.txt')
        if result == FileOperationResult.SUCCESS:
            print("Datei erfolgreich gelöscht.")
    """
    attempt = 0  # Zähler für die Anzahl der Versuche
    last_error = None  # Variable für den letzten Fehler

    while attempt < retries:
        try:
            os.remove(file_path)  # Versuche, die Datei zu löschen

            # Überprüfen, ob die Datei tatsächlich gelöscht wurde
            if not os.path.exists(file_path):
                return FileOperationResult.SUCCESS  # Erfolgreich gelöscht

        except FileNotFoundError:
            last_error = f"Die Datei '{file_path}' wurde nicht gefunden."  # Fehlerbeschreibung
            break  # Beende die Schleife, da die Datei nicht existiert
        except PermissionError:
            last_error = f"Keine Berechtigung, um die Datei '{file_path}' zu löschen."  # Fehlerbeschreibung
            break  # Beende die Schleife, da keine Berechtigung besteht
        except Exception as e:
            last_error = str(e)  # Speichere den Fehler
            attempt += 1  # Erhöhe den Versuchszähler
            time.sleep(delay_ms / 1000)  # Wartezeit in Sekunden

    # Rückgabe im Falle von Misserfolg nach allen Versuchen
    if last_error:
        app_logger.error(f"Datei '{file_path}' konnte nicht gelöscht werden: {last_error}")  # Protokolliere den letzten Fehler
        return FileOperationResult.UNKNOWN_ERROR  # Unbekannter Fehler
    else:
        app_logger.error(f"Datei '{file_path}' konnte nicht gelöscht werden.")
        return FileOperationResult.UNKNOWN_ERROR  # Rückgabe, wenn die Datei nicht gelöscht werden konnte


def sanitize_filename(filename: str) ->str:
    """
    Ersetzt ungültige Zeichen im Dateinamen durch Unterstriche.

    Parameter:
    filename (str): Der ursprüngliche Dateiname.

    Rückgabewert:
    str: Der bereinigte Dateiname, der nur gültige Zeichen enthält.
    """
    return re.sub(r'[<>:"/\\|?*]', '_', filename)


def format_datetime_stamp(datetime_stamp, format_string: str) -> str:
    """
    Formatiert einen Zeitstempel in das angegebene Format.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der formatiert werden soll.
    format_string (str): Das gewünschte Format für den Zeitstempel.

    Rückgabewert:
    str: Der formatierte Zeitstempel als String.
    """
    # Überprüfen, ob datetime_stamp ein datetime-Objekt ist
    if isinstance(datetime_stamp, datetime.datetime):
        dt = datetime_stamp  # Direkt verwenden, wenn es ein datetime-Objekt ist
    else:
        # Annehmen, dass es ein String ist und in ein datetime-Objekt umwandeln
        dt = datetime.datetime.strptime(datetime_stamp, "%Y-%m-%d %H:%M:%S")  # Format anpassen, falls nötig

    # Formatieren des datetime-Objekts gemäß dem bereitgestellten Format
    return dt.strftime(format_string)


def set_file_modification_date(file_path: str, new_date: str) -> FileOperationResult:
    """
    Setzt das Änderungsdatum einer Datei auf einen vorgegebenen Wert.

    Diese Funktion konvertiert das angegebene Änderungsdatum in einen Zeitstempel
    und verwendet die Windows-API, um das Änderungsdatum der angegebenen Datei zu ändern.
    Bei Erfolg wird ein Status von SUCCESS zurückgegeben. Bei Fehlern wird ein entsprechender
    Statuscode zurückgegeben.

    Parameter:
    file_path (str): Der Pfad zur Datei, deren Änderungsdatum geändert werden soll.
    new_date (str): Das neue Änderungsdatum im Format 'YYYY-MM-DD HH:MM:SS'.

    Rückgabewert:
    FileOperationResult: Ein Enum-Wert, der den Erfolg oder Fehler beschreibt.

    Beispiel:
        result = set_file_set_file_modification_date2('example.txt', '2023-11-08 12:00:00')
        if result == FileOperationResult.SUCCESS:
            print("Änderungsdatum erfolgreich gesetzt.")
    """

    try:
        # Konvertiere das Datum in einen Zeitstempel
        timestamp = time.mktime(datetime.datetime.strptime(new_date, '%Y-%m-%d %H:%M:%S').timetuple())

        # Überprüfe, ob die Datei existiert
        if not os.path.exists(file_path):
            return FileOperationResult.FILE_NOT_FOUND

        # Hole den aktuellen Änderungszeitstempel
        current_timestamp = os.path.getmtime(file_path)
        #print(f"Aktueller Änderungszeitstempel: {current_timestamp}")

        # Überprüfe, ob das Datum bereits übereinstimmt
        if current_timestamp == timestamp:
            #print("Der Änderungszeitstempel der Datei ist bereits korrekt.")
            return FileOperationResult.TIMESTAMP_MATCH

        # Setze das Änderungsdatum der Datei
        os.utime(file_path, (timestamp, timestamp))  # Setze das Änderungsdatum der Datei
        return FileOperationResult.SUCCESS  # Erfolgreich gesetzt

    except FileNotFoundError:
        app_logger.error(f"Die Datei '{file_path}' wurde nicht gefunden.")  # Protokolliere den Fehler
        return FileOperationResult.FILE_NOT_FOUND  # Datei nicht gefunden

    except PermissionError:
        app_logger.error(f"Keine Berechtigung, um das Änderungsdatum der Datei '{file_path}' zu ändern.")  # Protokolliere den Fehler
        return FileOperationResult.PERMISSION_DENIED  # Berechtigungsfehler

    except ValueError:
        app_logger.error(f"Ungültiges Datumsformat für '{new_date}'.")  # Protokolliere den Fehler
        return FileOperationResult.VALUE_ERROR  # Ungültiger Wert

    except Exception as e:
        app_logger.error(f"Allgemeiner Fehler beim Setzen des Änderungsdatums für '{file_path}': {str(e)}")  # Protokolliere den Fehler
        return FileOperationResult.UNKNOWN_ERROR  # Unbekannter Fehler


def set_file_creation_date(file_path: str, new_creation_date: str) -> FileOperationResult:
    """
    Setzt das Erstelldatum einer Datei auf einen vorgegebenen Wert.

    Diese Funktion konvertiert das angegebene Erstelldatum in einen Zeitstempel
    und verwendet die Windows-API, um das Erstelldatum der angegebenen Datei zu ändern.
    Bei Erfolg wird ein Status von SUCCESS zurückgegeben. Bei Fehlern wird ein entsprechender
    Statuscode zurückgegeben.

    Parameter:
    file_path (str): Der Pfad zur Datei, deren Erstelldatum geändert werden soll.
    new_creation_date (str): Das neue Erstelldatum im Format 'YYYY-MM-DD HH:MM:SS'.

    Rückgabewert:
    FileOperationResult: Ein Enum-Wert, der den Erfolg oder Fehler beschreibt.
    """

    try:
        # Konvertiere das Datum in einen Zeitstempel
        timestamp = time.mktime(time.strptime(new_creation_date, '%Y-%m-%d %H:%M:%S'))
        creation_time = pywintypes.Time(timestamp)  # Erstelle ein Zeitobjekt für die Windows-API

        # Überprüfe, ob die Datei existiert
        if not os.path.exists(file_path):
            return FileOperationResult.FILE_NOT_FOUND

        # Hole den aktuellen Erstellungszeitstempel
        current_timestamp = os.path.getctime(file_path)
        #print(f"Aktueller Erstellungszeitstempel: {current_timestamp}")

        # Überprüfe, ob das Datum bereits übereinstimmt
        if current_timestamp == timestamp:
            #print("Der Erstellungszeitstempel der Datei ist bereits korrekt.")
            return FileOperationResult.TIMESTAMP_MATCH

        # Verwende den Kontextmanager, um die Datei sicher zu öffnen
        with FileHandle(file_path) as handle:
            # Setze das Erstelldatum der Datei
            win32file.SetFileTime(handle, creation_time, None, None)

        return FileOperationResult.SUCCESS  # Erfolgreich gesetzt

    except FileNotFoundError:
        app_logger.error(f"Die Datei '{file_path}' wurde nicht gefunden.")
        return FileOperationResult.FILE_NOT_FOUND  # Datei nicht gefunden

    except PermissionError:
        app_logger.error(f"Keine Berechtigung, um das Erstelldatum der Datei '{file_path}' zu ändern.")
        return FileOperationResult.PERMISSION_DENIED  # Berechtigungsfehler

    except ValueError:
        app_logger.error(f"Ungültiges Datumsformat für '{new_creation_date}'.")
        return FileOperationResult.VALUE_ERROR  # Ungültiger Wert

    except Exception as e:
        app_logger.error(f"Allgemeiner Fehler beim Setzen des Erstelldatums für '{file_path}': {str(e)}")
        return FileOperationResult.UNKNOWN_ERROR  # Unbekannter Fehler


def delete_directory_contents(directory_path):
    """
    Löscht den gesamten Inhalt des angegebenen Verzeichnisses.

    Parameter:
    directory_path (str): Der Pfad des Verzeichnisses, dessen Inhalt gelöscht werden soll.

    Rückgabewert:
    str: Eine Bestätigung, dass der Inhalt erfolgreich gelöscht wurde.

    Wirft:
    OSError: Wenn das Löschen des Inhalts nicht erfolgreich ist.
    """
    if not os.path.isdir(directory_path):
        raise OSError(f"{directory_path} ist kein gültiges Verzeichnis.")

    try:
        for item in os.listdir(directory_path):
            item_path = os.path.join(directory_path, item)
            if os.path.isfile(item_path):
                os.remove(item_path)  # Löscht die Datei
            else:
                shutil.rmtree(item_path)  # Löscht das Verzeichnis rekursiv
        return "Inhalt erfolgreich gelöscht."
    except Exception as e:
        raise OSError(f"Fehler beim Löschen des Inhalts: {e}")

def copy_directory_contents(source_directory_path, target_directory_path):
    """
    Kopiert den gesamten Inhalt des angegebenen Quellverzeichnisses in das Zielverzeichnis.

    Parameter:
    source_directory_path (str): Der Pfad des Quellverzeichnisses.
    target_directory_path (str): Der Pfad des Zielverzeichnisses.

    Gibt:
    str: Eine Bestätigung, dass der Inhalt erfolgreich kopiert wurde.

    Wirft:
    OSError: Wenn das Kopieren des Inhalts nicht erfolgreich ist.
    """
    if not os.path.isdir(source_directory_path):
        raise OSError(f"{source_directory_path} ist kein gültiges Quellverzeichnis.")

    os.makedirs(target_directory_path, exist_ok=True)  # Erstellt das Zielverzeichnis, falls es nicht existiert

    try:
        for item in os.listdir(source_directory_path):
            s = os.path.join(source_directory_path, item)
            d = os.path.join(target_directory_path, item)
            if os.path.isdir(s):
                shutil.copytree(s, d, False, None)  # Kopiert Verzeichnisse rekursiv
            else:
                shutil.copy2(s, d)  # Kopiert Dateien
        return "Inhalt erfolgreich kopiert."
    except Exception as e:
        raise OSError(f"Fehler beim Kopieren des Inhalts: {e}")

