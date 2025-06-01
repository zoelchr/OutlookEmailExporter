"""
msg_handling.py

Dieses Modul enthält Funktionen zum Abrufen von Metadaten aus MSG-Dateien.
Es bietet Routinen, um verschiedene Informationen wie Absender, Empfänger,
Betreff und andere relevante Daten zu extrahieren.

Funktionen:
- get_msg_object(msg_file): Öffnet eine MSG-Datei, extrahiert relevante Daten und gibt sie als Dictionary zurück.
- create_log_file(base_name, directory, table_header): Erstellt ein Logfile im Excel-Format mit einem Zeitstempel im Namen.
- log_entry(log_file_path, entry): Fügt einen neuen Eintrag in das Logfile hinzu.
- convert_to_utc_naive(datetime_stamp): Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.
- format_datetime(datetime_stamp, format_string): Formatiert einen Zeitstempel in das angegebene Format.
- custom_sanitize_text(encoded_textstring): Bereinigt einen Textstring von unerwünschten Zeichen.
- truncate_filename_if_needed(file_path, max_length, truncation_marker): Kürzt den Dateinamen, wenn nötig.
- parse_sender_msg_file(msg_absender_str): Analysiert den Sender-String eines MSG-Files und extrahiert den Namen und die E-Mail-Adresse.
- load_known_senders(file_path): Lädt bekannte Sender aus einer CSV-Datei.
- reduce_thread_in_msg_message(email_text, max_older_emails): Reduziert die Anzahl der angehängten älteren E-Mails auf max_older_emails.

Verwendung:
Importiere die Funktionen aus diesem Modul in deinem Hauptprogramm oder anderen Modulen,
um auf die Metadaten von MSG-Dateien zuzugreifen.

Beispiel:
    from modules.msg_handling import get_msg_object
    msg_data = get_msg_object('example.msg')
"""

import extract_msg
import re
import os
import pandas as pd
from datetime import datetime
import logging
from enum import Enum

logger = logging.getLogger(__name__)

class MsgAccessStatus(Enum):
    SUCCESS = "Success"
    DATA_NOT_FOUND = "Data not found"
    FILE_NOT_FOUND = "File not found"
    PERMISSION_ERROR = "No permission to access"
    FILE_LOCKED = "File is locked by another process"
    ATTRIBUTE_ERROR = "Attribute error"
    UNICODE_DECODE_ERROR = "Unicode decode error"
    UNICODE_ENCODE_ERROR = "Unicode encode error"
    TYPE_ERROR = "Type error"
    VALUE_ERROR = "Value error"
    OTHER_ERROR = "Other error"
    UNKNOWN = "Unknown"
    NO_MESSAGE_FOUND = "No message found"
    NO_SENDER_FOUND = "No sender found"
    NO_RECIPIENT_FOUND = "No recipient found"
    SUBJECT_MISSING = "Subject missing"
    SENDER_MISSING = "Sender missing"
    DATE_MISSING = "Date missing"
    BODY_MISSING = "Body missing"
    ATTACHMENTS_MISSING = "Attachments missing"

def get_msg_object(msg_file: str) -> dict:
    """
    Öffnet eine MSG-Datei, extrahiert relevante Daten und gibt sie als Dictionary zurück.

    Diese Funktion versucht, ein MSG-Objekt aus der angegebenen MSG-Datei zu erstellen.
    Bei Erfolg werden die relevanten Informationen wie Betreff, Absender, Datum,
    Nachrichtentext und Anhänge extrahiert. Bei Fehlern wird der entsprechende Status
    in der Rückgabe angezeigt.

    Parameter:
    msg_file (str): Der Pfad zur MSG-Datei, die geöffnet werden soll.

    Rückgabewert:
    dict: Ein Dictionary mit den extrahierten Daten und einer Liste von Statuscodes:
        - "subject": Der Betreff der Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "sender": Der Absender der Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "date": Das Datum der gesendeten Nachricht oder "Unbekannt", wenn nicht vorhanden.
        - "body": Der Inhalt der Nachricht oder "Kein Inhalt verfügbar", wenn nicht vorhanden.
        - "attachments": Eine Liste der Dateinamen der Anhänge oder eine leere Liste, wenn keine vorhanden sind.
        - "status": Eine Liste von Statuscodes, die den Erfolg oder Fehler des Zugriffs beschreiben.
        - "signed": Boolean, ob die Nachricht signiert ist.
        - "encrypted": Boolean, ob die Nachricht verschlüsselt ist.
        - "reply_count": Anzahl der Antworten oder Weiterleitungen.
        - "has_defects": Boolean, ob die Nachricht Hinweise auf Defekte enthält.
    """

    # Initialisierung des Rückgabewerts mit Standardwerten
    msg_data = {
        "subject": "Unbekannt",
        "sender": "Unbekannt",
        "recipient": "Unbekannt",
        "date": "Unbekannt",
        "body": "Kein Inhalt verfügbar",
        "attachments": [],
        "status": [MsgAccessStatus.UNKNOWN],  # Liste mit Statuscodes
        "signed": False,
        "encrypted": False,
        "reply_count": 0,
        "has_defects": False
    }

    try:
        logger.debug(f"Öffne MSG-Datei: {msg_file}")  # Debugging-Ausgabe

        # Sicherstellen, dass die Datei mit `with` geöffnet und automatisch geschlossen wird
        with extract_msg.Message(msg_file) as msg_object:

            # Überprüfen, ob das MSG-Objekt erfolgreich erstellt wurde
            if msg_object is None:
                msg_data["status"].append(MsgAccessStatus.DATA_NOT_FOUND)
                logger.error(f"MSG-Datei konnte nicht verarbeitet werden: {msg_file}")
                return msg_data  # Sofort zurückgeben

            # Jedes Attribut separat absichern
            try:
                if msg_object.subject:
                    msg_data["subject"] = msg_object.subject
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Betreff erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.SUBJECT_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.sender:
                    msg_data["sender"] = msg_object.sender
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Sender erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.SENDER_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.recipients:
                    msg_data["recipient"] = msg_object.to
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Sender erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.NO_RECIPIENT_FOUND)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.date:
                    msg_data["date"] = msg_object.date
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Datum erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.DATE_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            try:
                if msg_object.body:
                    msg_data["body"] = msg_object.body
                    msg_data["status"] = [MsgAccessStatus.SUCCESS]  # Setze SUCCESS, wenn Body erfolgreich extrahiert
                else:
                    msg_data["status"].append(MsgAccessStatus.BODY_MISSING)
            except UnicodeDecodeError:
                msg_data["status"].append(MsgAccessStatus.UNICODE_DECODE_ERROR)
            except UnicodeEncodeError:
                msg_data["status"].append(MsgAccessStatus.UNICODE_ENCODE_ERROR)

            try:
                if msg_object.attachments:
                    msg_data["attachments"] = [att.longFilename or att.shortFilename or "unbenannt" for att in msg_object.attachments]
                    #msg_data["attachments"] = msg_object.attachments
                else:
                    msg_data["status"].append(MsgAccessStatus.ATTACHMENTS_MISSING)
            except AttributeError:
                msg_data["status"].append(MsgAccessStatus.ATTRIBUTE_ERROR)

            # Zusätzliche Informationen extrahieren
            msg_data["signed"] = hasattr(msg_object, 'signed') and msg_object.signed
            msg_data["encrypted"] = hasattr(msg_object, 'encrypted') and msg_object.encrypted
            msg_data["reply_count"] = getattr(msg_object, 'reply_count', 0)
            msg_data["has_defects"] = hasattr(msg_object, 'has_defects') and msg_object.has_defects

        logger.debug(f"MSG-Daten erfolgreich extrahiert: {msg_data}")  # Debugging-Ausgabe

    except FileNotFoundError:
        msg_data["status"] = [MsgAccessStatus.FILE_NOT_FOUND]
        logger.error(f"MSG-Datei nicht gefunden: {msg_file}")

    except PermissionError:
        msg_data["status"] = [MsgAccessStatus.PERMISSION_ERROR]
        logger.error(f"Keine Berechtigung, um die MSG-Datei zu öffnen: {msg_file}")

    except TypeError:
        msg_data["status"] = [MsgAccessStatus.TYPE_ERROR]
        logger.error(f"Falscher Datentyp in MSG-Datei: {msg_file}")

    except ValueError:
        msg_data["status"] = [MsgAccessStatus.VALUE_ERROR]
        logger.error(f"Ungültiger Wert in MSG-Datei: {msg_file}")

    except Exception as e:
        msg_data["status"] = [MsgAccessStatus.OTHER_ERROR]
        logger.error(f"Allgemeiner Fehler beim Öffnen der MSG-Datei: {str(e)}")

    return msg_data


def create_log_file(base_name, directory, table_header):
    """
    Erstellt ein Logfile im Excel-Format mit einem Zeitstempel im Namen.

    Parameter:
    base_name (str): Der Basisname des Logfiles.
    directory (str): Das Verzeichnis, in dem das Logfile gespeichert werden soll.

    Rückgabewert:
    str: Der Pfad zur erstellten Logdatei.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{base_name}_{timestamp}.xlsx"
    log_file_path = os.path.join(directory, log_file_name)

    # Leeres DataFrame mit den gewünschten Spalten erstellen
    df = pd.DataFrame(columns=table_header)

    try:
        df.to_excel(log_file_path, index=False)
        logger.debug(f"Logging Excel-Datei erfolgreich erstellt: {log_file_path}")  # Debugging-Ausgabe: Log-File
        return log_file_path
    except Exception as e:
        logger.error(f"Fehler beim Erstellen der Logdatei: {e}")  # Debugging-Ausgabe: Log-File
        raise OSError(f"Fehler beim Erstellen der Logdatei: {e}")

def create_log_file_neu(base_name, directory, table_header, sheet_name="Log"):
    """
    Erstellt ein Logfile im Excel-Format mit Zeitstempel und optionalem Sheetnamen.

    Parameter:
    base_name (str): Der Basisname der Logdatei.
    directory (str): Das Zielverzeichnis für die Datei.
    table_header (list): Die Spaltenüberschriften für die leere Tabelle.
    sheet_name (str): Der Name des Sheets (Standard: "Log").

    Rückgabewert:
    str: Der Pfad zur erstellten Logdatei.
    """
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file_name = f"{base_name}_{timestamp}.xlsx"
    log_file_path = os.path.join(directory, log_file_name)

    # Leeres DataFrame mit Header erstellen
    df = pd.DataFrame(columns=table_header)

    try:
        with pd.ExcelWriter(log_file_path, engine="openpyxl", mode="w") as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        logger.debug(f"Logging Excel-Datei erfolgreich erstellt: {log_file_path}")
        return log_file_path
    except Exception as e:
        logger.error(f"Fehler beim Erstellen der Logdatei: {e}")
        raise OSError(f"Fehler beim Erstellen der Logdatei: {e}")

def log_entry(log_file_path, entry):
    """
    Fügt einen neuen Eintrag in das Logfile hinzu.

    Parameter:
    log_file_path (str): Der Pfad zur Logdatei.
    entry (dict): Ein Dictionary mit den Werten für die Logzeile.

    Rückgabewert:
    None
    """
    df = pd.read_excel(log_file_path)

    # Erstellen eines DataFrames aus dem Eintrag
    new_entry_df = pd.DataFrame([entry])

    # Überprüfen, ob new_entry_df leer ist oder nur NA-Werte enthält
    if not new_entry_df.empty and not new_entry_df.isnull().all(axis=1).any():
        # Logdatei laden oder erstellen
        if os.path.exists(log_file_path):
            with pd.ExcelFile(log_file_path) as xls:  # Verwende einen with-Block
                df = pd.read_excel(xls)
                # Nur nicht-leere DataFrames zusammenführen
                if not df.empty and not df.isna().all(axis=None):
                     if not new_entry_df.empty and not new_entry_df.isna().all(axis=None):
                        df = pd.concat([df, new_entry_df], ignore_index=True)
                else:
                    df = new_entry_df  # Neue Logdatei erstellen

        # Speichern des aktualisierten DataFrames in die Logdatei
        df.to_excel(log_file_path, index=False)

def log_entry_neu(log_file_path, entry, sheet_name="Log"):
    """
    Fügt einen oder mehrere Einträge in das Logfile (Excel) hinzu.

    Parameter:
    - log_file_path (str): Der Pfad zur Excel-Logdatei.
    - entry (dict | list[dict]): Ein einzelner oder mehrere Logeinträge.
    - sheet_name (str): Der Name des Sheets (Standard: "Log").

    Rückgabewert:
    - None
    """

    # entry in DataFrame umwandeln (egal ob dict oder Liste)
    if isinstance(entry, dict):
        new_entry_df = pd.DataFrame([entry])
    elif isinstance(entry, list) and all(isinstance(e, dict) for e in entry):
        new_entry_df = pd.DataFrame(entry)
    else:
        print("Ungültiger Eintragstyp – erwartet dict oder Liste von dicts.")
        return

    # Leere oder nutzlose Einträge überspringen
    if new_entry_df.empty or new_entry_df.isnull().all(axis=1).all():
        print("Die übergebenen Daten sind leer oder vollständig ungültig – nichts gespeichert.")
        return

    # Versuche bestehende Datei und Sheet zu lesen
    if os.path.exists(log_file_path):
        try:
            with pd.ExcelFile(log_file_path) as xls:
                if sheet_name in xls.sheet_names:
                    df_alt = pd.read_excel(xls, sheet_name=sheet_name)
                else:
                    df_alt = pd.DataFrame()
        except Exception as e:
            print(f"Fehler beim Lesen der Excel-Datei: {e}")
            df_alt = pd.DataFrame()
    else:
        df_alt = pd.DataFrame()

    # Neue Daten anhängen oder neu starten
    if not df_alt.empty and not df_alt.isna().all(axis=None):
        df_neu = pd.concat([df_alt, new_entry_df], ignore_index=True)
    else:
        df_neu = new_entry_df

    # Schreiben (bestehendes Sheet ersetzen)
    try:
        with pd.ExcelWriter(log_file_path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            df_neu.to_excel(writer, sheet_name=sheet_name, index=False)
    except FileNotFoundError:
        with pd.ExcelWriter(log_file_path, engine="openpyxl", mode="w") as writer:
            df_neu.to_excel(writer, sheet_name=sheet_name, index=False)

def convert_to_utc_naive(datetime_stamp):
    """
    Konvertiert einen Zeitstempel in ein UTC-naives Datetime-Objekt.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der konvertiert werden soll.

    Rückgabewert:
    datetime: Ein UTC-naives Datetime-Objekt.
    """
    try:
        if datetime_stamp.tzinfo is not None:
            new_datetime_stamp = datetime_stamp.replace(tzinfo=None)
            logger.debug(f"Konvertierter Zeitstempel in ein UTC-naives Datetime-Objekt: {new_datetime_stamp}")  # Debugging-Ausgabe: Log-File
            return datetime_stamp.replace(tzinfo=None)  # Entfernen der Zeitzone

        logger.error(f"Kein Zeitstempel zum Konvertieren vorhanden.")  # Debugging-Ausgabe: Log-File
        return datetime_stamp

    except Exception as e:
        print(f"Fehler beim Konvertieren des Zeitstempels: {str(e)}") # Debugging-Ausgabe: Console
        logger.error(f"Fehler beim Konvertieren des Zeitstempels: {str(e)}") # Debugging-Ausgabe: Log-File
        return datetime_stamp


def format_datetime(datetime_stamp, format_string):
    """
    Formatiert einen Zeitstempel in das angegebene Format.

    Parameter:
    datetime_stamp (datetime): Der Zeitstempel, der formatiert werden soll.
    format_string (str): Das gewünschte Format für den Zeitstempel.

    Rückgabewert:
    str: Der formatierte Zeitstempel als String.

    Beispiel:
        formatted_time = format_datetime(datetime.now(), "%Y-%m-%d %H:%M:%S")
    """
    if datetime_stamp is None:
        logger.error(f"Der Zeitstempel darf nicht None sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("Der Zeitstempel darf nicht None sein.")

    if not isinstance(datetime_stamp, datetime):
        logger.error(f"Ungültiger Zeitstempel: erwartet datetime, erhalten {type(datetime_stamp).__name__}.")  # Debugging-Ausgabe: Log-File
        raise ValueError(f"Ungültiger Zeitstempel: erwartet datetime, erhalten {type(datetime_stamp).__name__}.")

    return datetime_stamp.strftime(format_string)


def custom_sanitize_text(encoded_textstring):
    """
    Bereinigt einen Textstring, indem unerwünschte Zeichen ersetzt und Formatierungen angepasst werden.

    Diese Funktion führt mehrere Schritte zur Bereinigung des Eingabetextes durch:
    1. Ersetzt mehrere aufeinanderfolgende Leerzeichen durch ein einzelnes Leerzeichen.
    2. Entfernt Leerzeichen am Ende des Textes.
    3. Ersetzt spezifische unerwünschte Zeichenfolgen durch definierte Alternativen.
    4. Ersetzt unerwünschte Zeichen durch sichere Alternativen, um sicherzustellen, dass der Text als Dateiname verwendet werden kann.

    Parameter:
    encoded_textstring (str): Der ursprüngliche Textstring, der bereinigt werden soll.

    Rückgabewert:
    str: Der bereinigte Textstring, der als gültiger Dateiname verwendet werden kann,
         wobei unerwünschte Zeichen entfernt oder ersetzt wurden.

    Beispiel:
        sanitized_string = custom_sanitize_text("Beispiel: ungültige Zeichen / \\ * ? < > |")
    """
    # Ersetze mehrere aufeinanderfolgende Leerzeichen durch ein einzelnes Leerzeichen
    encoded_textstring = re.sub(r'\s+', ' ', encoded_textstring)
    # Entferne Leerzeichen am Ende
    encoded_textstring = encoded_textstring.rstrip()

    # Ersetze spezielle Zeichenfolgen
    encoded_textstring = encoded_textstring.replace("_-_", "-")
    encoded_textstring = encoded_textstring.replace(" - ", "-")
    encoded_textstring = encoded_textstring.replace("._", "_")
    encoded_textstring = encoded_textstring.replace("_.", "_")
    encoded_textstring = encoded_textstring.replace(" .", "_")
    encoded_textstring = encoded_textstring.replace(". ", "_")
    encoded_textstring = encoded_textstring.replace(" / ", "_")
    encoded_textstring = encoded_textstring.replace(" & ", "_")
    encoded_textstring = encoded_textstring.replace("; ", "_")
    encoded_textstring = encoded_textstring.replace("/ ", "_")
    encoded_textstring = encoded_textstring.replace(" | ", "_")

    replacements = {
        " ": "_",
        "#": "_",
        "%": "_",
        "&": "_",
        "*": "-",
        "{": "-",
        "}": "-",
        "\\": "-",
        ":": "",
        "<": "-",
        ">": "-",
        "?": "-",
        "/": "_",
        "|": "_",
        "\"": "",
        "ä": "ae",
        "Ä": "Ae",
        "ö": "oe",
        "Ö": "Oe",
        "ü": "ue",
        "Ü": "Ue",
        "ß": "ss",
        "é": "e",
        ",": "",
        "!": "",
        "'": "_",
        ";": "_",
        "“": "",
        "„": ""
    }

    for old_char, new_char in replacements.items():
        encoded_textstring = encoded_textstring.replace(old_char, new_char)

    return encoded_textstring


def truncate_filename_if_needed(file_path, max_length, truncation_marker):
    """
    Kürzt den Dateinamen, wenn der gesamte Pfad die maximal zulässige Länge überschreitet.

    Diese Funktion überprüft die Länge des angegebenen Dateipfads und vergleicht sie mit der maximalen
    zulässigen Länge. Wenn der Pfad diese Länge überschreitet, wird der Dateiname so gekürzt, dass er
    zusammen mit dem Verzeichnispfad die maximale Länge nicht überschreitet. Am Ende des gekürzten
    Dateinamens wird ein Truncation Marker hinzugefügt, um anzuzeigen, dass der Name gekürzt wurde.

    Parameter:
    file_path (str): Der vollständige Dateipfad, der überprüft und möglicherweise gekürzt werden soll.
    max_length (int): Die maximal zulässige Länge des gesamten Dateipfads.
    truncation_marker (str): Die Zeichenkette, die verwendet wird, um das Kürzen anzuzeigen (z.B. "<>").

    Rückgabewert:
    str: Der möglicherweise gekürzte Dateipfad, der die maximal zulässige Länge nicht überschreitet.

    Beispiel:
        truncated_path = truncate_filename_if_needed("D:/Dev/pycharm/MSGFileRenamer/modules/very_long_filename_that_exceeds_the_limit.txt", 50, "...")
    """
    if file_path is None:
        logger.error(f"file_path darf nicht None sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("file_path darf nicht None sein.")

    if not isinstance(max_length, int) or max_length <= 0:
        logger.error(f"max_length muss eine positive Ganzzahl sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("max_length muss eine positive Ganzzahl sein.")

    if not truncation_marker:
        logger.error(f"truncation_marker darf nicht leer sein.")  # Debugging-Ausgabe: Log-File
        raise ValueError("truncation_marker darf nicht leer sein.")

    if len(file_path) > max_length:
        # Berechne die maximale Länge für den Dateinamen
        path_length = len(os.path.dirname(file_path))
        max_filename_length = max_length - path_length - len(truncation_marker) - 1  # -1 für den Schrägstrich

        # Extrahiere den Dateinamen
        filename = os.path.basename(file_path)

        if len(filename) > max_filename_length:
            # Kürze den Dateinamen und füge den Truncation Marker hinzu
            truncated_filename = filename[:max_filename_length] + truncation_marker
            return os.path.join(os.path.dirname(file_path), truncated_filename)

    return file_path

def parse_sender_msg_file(msg_absender_str: str) -> dict:
    """
    Analysiert den Sender-String eines MSG-Files und extrahiert den Namen und die E-Mail-Adresse.

    Parameter:
    sender (str): Der Sender-String.

    Rückgabewert:
    dict: Ein Dictionary mit 'sender_name', 'sender_email' und 'contains_sender_email'.
    """

    # Regulärer Ausdruck für die E-Mail-Adresse
    email_pattern = r'<(.*?)>'
    email_match = re.search(email_pattern, msg_absender_str)

    if email_match:
        sender_email = email_match.group(1)
        contains_sender_email = True
        logger.debug(f"Im Absender der MSG-Datei ist folgende Email enthalten: {sender_email}")  # Debugging-Ausgabe: Log-File
    else:
        sender_email = ""
        contains_sender_email = False
        logger.debug(f"Im Absender der MSG-Datei ist keine Email enthalten: {msg_absender_str}")  # Debugging-Ausgabe: Log-File

    # Entferne die E-Mail-Adresse aus dem Sender-String
    sender_name = re.sub(email_pattern, '', msg_absender_str).strip()
    # Entferne Anführungszeichen aus dem Sender-String
    sender_name = sender_name.replace("\"", '')

    return {
        "sender_name": sender_name,
        "sender_email": sender_email,
        "contains_sender_email": contains_sender_email
    }


def load_known_senders(file_path):
    """
    Lädt bekannte Sender aus einer CSV-Datei.

    Parameter:
    file_path (str): Der Pfad zur CSV-Datei.

    Rückgabewert:
    DataFrame: Ein DataFrame mit den bekannten Sendern.
    """
    return pd.read_csv(file_path)


def reduce_thread_in_msg_message(email_text, max_older_emails=2) -> dict:
    """
    Reduziert die Anzahl der angehängten älteren E-Mails auf max_older_emails.
    Ältere E-Mails werden anhand der typischen Kopfzeilen (Von, Gesendet, An, Cc, Betreff) erkannt.

    :param email_text: Der vollständige Text der E-Mail
    :param max_older_emails: die maximale Anzahl an beizubehaltenden alten E-Mails
    :return: ein Dictionary mit dem bereinigten E-Mail-Text und der Anzahl der gelöschten alten E-Mails
    """

    # Regulärer Ausdruck für eine typische E-Mail-Kopfzeile
    email_header_pattern = re.compile(
        r"(Von: .+?Betreff: .+?)(\n\n|\r\n\r\n)", re.DOTALL | re.IGNORECASE
    )

    # Alle gefundenen älteren E-Mails identifizieren
    older_emails_text = email_header_pattern.split(email_text)

    # Anzahl der gefundenen älteren E-Mails
    total_older_emails = (len(older_emails_text) // 3) - 1  # Abzüglich der ursprünglichen Nachricht

    if total_older_emails <= max_older_emails:
        return {"new_email_text": email_text, "deleted_count": 0}  # Keine Kürzung nötig

    # Behalten der neuesten E-Mail + der maximal erlaubten Anzahl alter E-Mails
    new_email_text = older_emails_text[0]  # Original-Nachricht ohne Anhang
    for i in range(1, min((max_older_emails * 3) + 1, len(older_emails_text)), 3):
        new_email_text += older_emails_text[i] + older_emails_text[i + 1]

    # Berechnung der Anzahl der gelöschten E-Mails
    deleted_count = total_older_emails - max_older_emails

    # Hinzufügen eines Hinweises für entfernte ältere E-Mails
    new_email_text += "\n\n--- deleted_count ältere E-Mails wurden entfernt. Vollständige E-Mail-Kette im Projektpostfach einsehbar. ---\n"

    return {"new_email_text": new_email_text, "deleted_count": deleted_count}
