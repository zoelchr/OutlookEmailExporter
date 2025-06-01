# -*- coding: utf-8 -*-
"""
msg_generate_new_filename.py

Dieses Modul enthält Funktionen zur Generierung eines neuen Dateinamens für MSG-Dateien basierend auf deren Metadaten.
Es extrahiert Informationen wie den Absender, das Versanddatum und den Betreff der Nachricht, um einen eindeutigen und informativen Dateinamen zu erstellen.

Funktionen:
- generate_new_msg_filename(msg_path_and_filename, max_path_length=260): Generiert einen neuen Dateinamen für eine MSG-Datei basierend auf Metadaten wie Absender, Versanddatum und Betreff. Kürzt den Dateinamen, falls er die maximale Pfadlänge überschreitet.

Verwendung:
Importieren Sie dieses Modul in Ihr Skript, um neue Dateinamen für MSG-Dateien zu generieren, die auf den Metadaten der Dateien basieren.
"""

import os
import logging
from datetime import datetime  # Stellen Sie sicher, dass nur die Klasse datetime importiert wird
from modules.msg_handling import parse_sender_msg_file, \
    load_known_senders, convert_to_utc_naive, format_datetime, \
    custom_sanitize_text, truncate_filename_if_needed, MsgAccessStatus, get_msg_object
from dataclasses import dataclass

@dataclass
class MsgFilenameResult:
    """
    MsgFilenameResult

    Diese Datenklasse speichert die Ergebnisse der Dateinamensgenerierung für MSG-Dateien. Sie enthält alle relevanten Informationen, die während des Prozesses extrahiert und verarbeitet wurden.

    Attribute:
    - datetime_stamp: Der ursprüngliche Zeitstempel des Versanddatums als datetime-Objekt.
    - formatted_timestamp: Der formatierte Zeitstempel als String, basierend auf dem gewünschten Format.
    - sender_name: Der Name des Absenders der Nachricht.
    - sender_email: Die E-Mail-Adresse des Absenders.
    - msg_subject: Der ursprüngliche Betreff der Nachricht.
    - msg_subject_sanitized: Der bereinigte Betreff, bei dem unerwünschte Zeichen entfernt wurden.
    - new_msg_filename: Der neu generierte Dateiname für die MSG-Datei.
    - new_truncated_msg_filename: Der gekürzte Dateiname, falls der ursprüngliche Dateiname die maximale Pfadlänge überschreitet.
    - is_msg_filename_truncated: Ein boolescher Wert, der angibt, ob der Dateiname gekürzt wurde.

    Verwendung:
    Diese Klasse wird verwendet, um die Ergebnisse der Funktion `generate_new_msg_filename` zu speichern und zurückzugeben.
    """
    datetime_stamp: datetime
    formatted_timestamp: str
    sender_name: str
    sender_email: str
    msg_subject: str
    msg_subject_sanitized: str
    new_msg_filename: str
    new_truncated_msg_filename: str
    is_msg_filename_truncated: bool

logger = logging.getLogger(__name__)

# Liste der bekannten Email-Absender aus einer CSV-Datei
LIST_OF_KNOWN_SENDERS = r'.\config\known_senders_private.csv'

PRINT_RESULT = False

def generate_new_msg_filename(msg_path_and_filename, use_list_of_known_senders=False, file_list_of_known_senders=LIST_OF_KNOWN_SENDERS, max_console_output=False, max_path_length=260):
    """
    generate_new_msg_filename(msg_path_and_filename, max_path_length=260)

    Diese Funktion generiert einen neuen Dateinamen für eine MSG-Datei basierend auf deren Metadaten. Der neue Dateiname wird aus dem Versanddatum, der Absender-E-Mail und dem Betreff der Nachricht zusammengesetzt. Falls der resultierende Dateipfad die maximale Pfadlänge überschreitet, wird der Dateiname entsprechend gekürzt.

    Parameter:
    - msg_path_and_filename: Der vollständige Pfad zur MSG-Datei, für die ein neuer Dateiname generiert werden soll.
    - max_path_length: Die maximale Länge des Dateipfads. Standardmäßig auf 260 Zeichen gesetzt.

    Rückgabe:
    - Ein MsgFilenameResult-Objekt, das Informationen wie den Zeitstempel, den formatierten Zeitstempel, den Absendernamen, die Absender-E-Mail, den Betreff, den bereinigten Betreff, den neuen Dateinamen, den gekürzten Dateinamen (falls erforderlich) und einen Indikator, ob der Dateiname gekürzt wurde, enthält.

    Verwendung:
    Importieren Sie diese Funktion in Ihr Skript, um neue Dateinamen für MSG-Dateien zu generieren, die auf den Metadaten der Dateien basieren.
    """
    if max_console_output: print(f"\t*************************************************************")
    if max_console_output: print(f"\t* Versuche einen neuen Namen für die MSG-Datei zu generieren.")
    if max_console_output: print(f"\t*************************************************************")

    format_string = "%Y%m%d-%Huhr%M"  # Beispiel für das gewünschte Format für Zeitstempel

    # 0. Schritt: Laden der bekannten Sender aus der Tabelle der bekannten Email-Absender, wenn use_list_of_known_senders ist True
    if use_list_of_known_senders:
        logger.debug(f"\tSchritt 0: Versuche Einlesen Liste bekannter Email-Absender aus CSV-Datei: {file_list_of_known_senders}'")  # Debugging-Ausgabe: Log-File
        # Prüfen, ob die Datei existiert und lesbar ist
        if os.path.exists(file_list_of_known_senders) and os.access(file_list_of_known_senders, os.R_OK):
            if max_console_output: print(f"\tSchritt 0: Die Tabelle der bekannten Email-Absender '{file_list_of_known_senders}' ist zugänglich und lesbar.")
            logger.debug(f"Schritt 0: Die Tabelle der bekannten Email-Absender '{file_list_of_known_senders}' ist zugänglich und lesbar.")  # Debugging-Ausgabe: Log-File

            known_senders_df = load_known_senders(file_list_of_known_senders)  # Laden der bekannten Sender aus der CSV-Datei als Dataframe
            logger.debug(f"Schritt 0: Liste der bekannten Email-Absender: {known_senders_df}'")  # Debugging-Ausgabe: Log-File
            exist_csv_file = True
        else:
            if max_console_output: print(f"\tSchritt 0: Die Tabelle der bekannten Email-Absender '{file_list_of_known_senders}' ist nicht zugänglich bzw. nicht lesbar.")
            exist_csv_file = False
    else:
        if max_console_output: print(f"\tSchritt 0: Die Tabelle der bekannten Email-Absender wird nicht genutzt.")
        exist_csv_file = False

    # Auslesen des msg-Objektes
    msg_object = get_msg_object(msg_path_and_filename)

    # 1. Schritt: Absender-String aus der MSG-Datei abrufen mit alternativer Methode
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object["status"]:
        found_msg_sender_string = msg_object["sender"]  # Absender extrahieren
        if max_console_output: print(f"\tSchritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'") # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 1: In MSG-Datei gefundener Absender-String: {found_msg_sender_string}'")  # Debugging-Ausgabe: Log-File
    else:
        found_msg_sender_string = ""
        if max_console_output: print(f"\tSchritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Console
        logger.warning(f"Schritt 1: In MSG-Datei keinen Absender-String gefunden.")  # Debugging-Ausgabe: Log-File

    # 2. Schritt: Absender-Email aus dem gefundenen Absender-String mithilfe einer Regex-Methode extrahieren
    parsed_sender_email = {"sender_name": "", "sender_email": "", "contains_sender_email": False} # Defaultwerte für parsed_sender_email setzen

    # Wenn der Absender-String aus der MSG-Datei erfolgreich ausgelesen wurde, dann wird die Absender-Email aus dem Absender-String extrahiert
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SENDER_MISSING not in msg_object["status"]:
        parsed_sender_email = parse_sender_msg_file(found_msg_sender_string)
        if max_console_output: print(f"\tSchritt 2: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 2: Absender-Email in Absender-String der MSG-Datei gefunden: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Log-File
    else:
        if max_console_output: print(f"\tSchritt 2: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt.")
        logger.debug(f"Schritt 2: Absender-String der MSG-Datei ist fehlerhaft oder unbekannt.")

    #  3. Schritt: Wenn im 2. Schritt keine Absender-Email gefunden wurde, dann in der Tabelle der bekannten Absender-Emails nachsehen, ob der Absendername enthalten ist
    if (not parsed_sender_email["contains_sender_email"]) and (use_list_of_known_senders):

        if exist_csv_file:
            known_sender_row = known_senders_df[known_senders_df['sender_name'].str.contains(parsed_sender_email["sender_name"], na=False, regex=False)]

            if not known_sender_row.empty:
                parsed_sender_email["sender_email"] = known_sender_row.iloc[0]["sender_email"]
                parsed_sender_email["contains_sender_email"] = True
                if max_console_output: print(f"\tSchritt 3: In Tabelle gefundene Absender-Email: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Console
                logger.debug(f"Schritt 3: In Tabelle gefundene Absender-Email: '{parsed_sender_email['sender_email']}'")  # Debugging-Ausgabe: Log-File
            else:
                parsed_sender_email["contains_sender_email"] = False
                if max_console_output: print(f"\tSchritt 3: In Tabelle keine Absender-Email für folgenden Absender-String gefunden: '{found_msg_sender_string}'")  # Debugging-Ausgabe: Console
                logger.warning(f"Schritt 3: In Tabelle keine Absender-Email für folgenden Absender-String gefunden: '{found_msg_sender_string}'")  # Debugging-Ausgabe: Log-File
        else:
            if max_console_output: print(f"\tSchritt 3: Kein Suchen in der Tabelle der bekannten Email-Absender möglich: '{file_list_of_known_senders}'")  # Debugging-Ausgabe: Console
            logger.debug(f"Schritt 3: Kein Suchen in Tabelle der bekannten Email-Absender möglich: '{file_list_of_known_senders}'")  # Debugging-Ausgabe: Log-File

    else:
        if max_console_output: print(f"\tSchritt 3: Kein Nachschlagen in der Tabelle der bekannten Email-Absender erforderlich bzw. gewünscht.")  # Debugging-Ausgabe: Console
        logger.warning(f"Kein Nachschlagen in der Tabelle der bekannten Email-Absender erforderlich bzw. gewünscht.")

    # 4. Schritt: Versanddatum abrufen und konvertieren
    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.DATE_MISSING not in msg_object["status"]:
        datetime_stamp = msg_object['date']
        datetime_stamp = convert_to_utc_naive(datetime_stamp)  # Sicherstellen, dass der Zeitstempel zeitzonenunabhängig ist
        if max_console_output: print(f"\tSchritt 4: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4: Versanddatum abrufen und konvertieren: '{datetime_stamp}'")  # Debugging-Ausgabe: Log-File

        # 4a. Schritt: Formatiertes Versanddatum ermitteln
        try: 
            formatted_timestamp = format_datetime(datetime_stamp, format_string)  # Formatieren des Zeitstempels
        except ValueError as e:
            datetime_stamp = ""
            formatted_timestamp = ""
            if max_console_output: print(f"\tSchritt 4: Kein Versanddatum gefunden, wegen Fehler: '{e}'")  # Debugging-Ausgabe: Console
            logger.debug(f"Schritt 4: Kein Versanddatum gefunden, wegen Fehler: '{e}'")  # Debugging-Ausgabe: Log-File

        if max_console_output: print(f"\tSchritt 4a: Formatiertes Versanddatum: '{formatted_timestamp}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4a: Formatiertes Versanddatum: '{formatted_timestamp}'")  # Debugging-Ausgabe: Log-File
    else:
        datetime_stamp = ""
        formatted_timestamp = ""
        if max_console_output: print(f"\tSchritt 4: Kein Versanddatum gefunden: '{msg_object['status']}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 4: Kein Versanddatum gefunden: '{msg_object['status']}'")  # Debugging-Ausgabe: Log-File

    if MsgAccessStatus.SUCCESS in msg_object["status"] and MsgAccessStatus.SUBJECT_MISSING not in msg_object["status"]:
        msg_subject = msg_object["subject"]
        if max_console_output: print(f"\tSchritt 5: Ermittelter Betreff: '{msg_subject}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 5: Betreff ermitteln: '{msg_subject}'")  # Debugging-Ausgabe: Log-File

        # 6. Schritt: Betreff bereinigen
        msg_subject_sanitized = custom_sanitize_text(msg_subject)  # Betreff bereinigen
        if max_console_output: print(f"\tSchritt 6: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 6: Bereinigten Betreff ermitteln: '{msg_subject_sanitized}'")  # Debugging-Ausgabe: Log-File

    else:
        msg_subject = ""
        msg_subject_sanitized = ""

    # 7. Schritt: Neuen Namen der Datei festlegen
    new_msg_filename = f"{formatted_timestamp}_{parsed_sender_email['sender_email']}_{msg_subject_sanitized}.msg"
    msg_pathname = os.path.dirname(msg_path_and_filename)  # Verzeichnisname der MSG-Datei
    if max_console_output: print(f"\tSchritt 7: Neuer Dateiname: '{new_msg_filename}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 7: Neuer Dateiname: '{new_msg_filename}'")  # Debugging-Ausgabe: Log-File

    new_msg_path_and_filename = os.path.join(msg_pathname, new_msg_filename)  # Neuer absoluter Dateipfad
    if max_console_output: print(f"\tSchritt 8: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Console
    logger.debug(f"Schritt 8: Neuer absoluter Dateiname: '{new_msg_path_and_filename}'")  # Debugging-Ausgabe: Log-File

    # 9. Schritt: Kürzen des Dateinamens, falls nötig
    if len(new_msg_path_and_filename) > max_path_length:
        new_truncated_msg_path_and_filename = truncate_filename_if_needed(new_msg_path_and_filename, max_path_length, "...msg")
        new_truncated_msg_filename = os.path.basename(new_truncated_msg_path_and_filename)
        is_msg_filename_truncated = True
        if max_console_output: print(f"\tSchritt 9: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 9: Neuer gekürzter Dateiname: '{new_truncated_msg_filename}'")  # Debugging-Ausgabe: Log-File
    else:
        new_truncated_msg_path_and_filename = new_msg_path_and_filename
        new_truncated_msg_filename = new_msg_filename
        is_msg_filename_truncated = False
        if max_console_output: print(f"\tSchritt 9: Kein Kürzen des Dateinamens erforderlich.")  # Debugging-Ausgabe: Console
        logger.debug(f"Schritt 9: Kein Kürzen des Dateinamens erforderlich.")  # Debugging-Ausgabe: Log-File

    # Ausgabe aller Informationen
    if PRINT_RESULT:
        print(f"\n\t'generate_new_msg_filename' - Ausgabe aller Informationen")
        print(f"\tVollständiger Pfad der MSG-Datei: {msg_path_and_filename}")
        print(f"\tEnthält Email-Absender: {parsed_sender_email['contains_sender_email']}")
        print(f"\tGefundener Email-Absender: {parsed_sender_email['sender_email']}")
        print(f"\tExtrahierter Versandzeitpunkt: {datetime_stamp}")
        print(f"\tFormatierter Versandzeitpunkt: {formatted_timestamp}")
        print(f"\tExtrahierter Betreff: {msg_subject}")
        print(f"\tBereinigter Betreff: {msg_subject_sanitized}")
        print(f"\tNeuer Dateiname: {new_msg_filename}")
        print(f"\tKürzung Dateiname erforderlich: {is_msg_filename_truncated}")
        print(f"\tNeuer gekürzter Dateiname: {new_truncated_msg_filename}\n")

    # Rückgabe der gewünschten Informationen als MsgFilenameResult
    return MsgFilenameResult(
        datetime_stamp=datetime_stamp,
        formatted_timestamp=formatted_timestamp,
        sender_name=parsed_sender_email["sender_name"],
        sender_email=parsed_sender_email["sender_email"],
        msg_subject=msg_subject,
        msg_subject_sanitized=msg_subject_sanitized,
        new_msg_filename=new_msg_filename,
        new_truncated_msg_filename=new_truncated_msg_filename,
        is_msg_filename_truncated=is_msg_filename_truncated
    )