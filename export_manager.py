import os
import re
import logging

from PySide6.QtWidgets import QMessageBox
import win32com.client
from dotenv import load_dotenv
import time

from modules.msg_generate_new_filename import generate_new_msg_filename
from utils.file_handling import rename_file, FileOperationResult
from config import KNOWNSENDER_FILE, TEST_MODE

# Erstellen eines Loggers für Protokollierung von Ereignissen und Fehlern
app_logger = logging.getLogger(__name__)

class ExportManager:
    def __init__(self, email_table_model, export_directory):
        """
        Initialisiert den Exporter mit einem Tabellenmodell und einem Exportverzeichnis.

        :param email_table_model: Instanz von EmailTableModel.
        :param export_directory: Ausgewähltes Exportziel (z. B. Verzeichnis).
        """
        self.email_table_model = email_table_model
        self.export_directory = export_directory

        app_logger.debug(f"ExportManager mit Tabellenmodell und Exportverzeichnis {export_directory} initialisiert.")


    def export_emails(self, export_type, change_filename, change_filedate, overwrite_file):
        """
        Exportiert die ausgewählten Emails in das Zielverzeichnis.
        """
        def show_error_dialog(message):
            """
            Zeigt eine Fehler-Dialogbox mit der angegebenen Nachricht an.

            :param message: Die anzuzeigende Fehlermeldung.
            """
            error_dialog = QMessageBox()
            error_dialog.setIcon(QMessageBox.Critical)
            error_dialog.setWindowTitle("Fehler")
            error_dialog.setText(message)
            error_dialog.exec()

        # 1. Überprüfen, ob das Exportverzeichnis existiert
        if not os.path.exists(self.export_directory):
            error_message = f"Das Exportverzeichnis existiert nicht: \n{self.export_directory}"
            app_logger.error(error_message)
            show_error_dialog(error_message)
            app_logger.debug(f"Das Exportverzeichnis existiert nicht: \n{self.export_directory}")
            return 0

        # 2. Überprüfen, ob Schreibrechte im Exportverzeichnis vorhanden sind
        if not os.access(self.export_directory, os.W_OK):
            error_message = f"Keine Schreibrechte im Exportverzeichnis: \n{self.export_directory}"
            app_logger.error(error_message)
            show_error_dialog(error_message)
            app_logger.debug(f"Keine Schreibrechte im Exportverzeichnis: \n{self.export_directory}")
            return 0

        app_logger.debug(f"Exportverzeichnis geprüft: {self.export_directory} (existiert und ist beschreibbar).")

        # 3. Abrufen der ausgewählten Emails
        selected_emails = self.email_table_model.get_selected_emails()
        app_logger.debug(f"Anzahl ausgewählter Emails: {len(selected_emails)}")

        if not selected_emails:
            error_message = f"Keine Emails zum Exportieren vorhanden."
            app_logger.error(error_message)
            show_error_dialog(error_message)
            app_logger.debug(f"Keine Emails zum Exportieren vorhanden.")
            return 0

        # Alle ausgewählten Emails abarbeiten
        for email in selected_emails:
            print(f"- {email}")

            # Speichern der Nachricht als MSG über Outlook mit einem vereinfachten Namen
            if not TEST_MODE:
                path_and_file_name = self.save_as_msg(email, test_mode=False)
            else:
                path_and_file_name = self.save_as_msg(email, test_mode=True)
                app_logger.warning(f"MSG-Date wegen Test-Modus nicht gespeichert.")

            app_logger.debug(f"Initialer absoluter Dateiname: {path_and_file_name}.")

            # Neuen Dateinamen erzeugen
            new_msg_filename_collection = generate_new_msg_filename(
                path_and_file_name,
                use_list_of_known_senders=True,
                file_list_of_known_senders=KNOWNSENDER_FILE
            )
            new_file_name = new_msg_filename_collection.new_msg_filename
            app_logger.debug(f"Neuer Dateiname: {new_file_name}.")

            if not new_file_name:  # Wenn new_file_name leer oder None ist
                app_logger.debug(f"Kein gültiger Dateiname gefunden für E-Mail: {email.subject}. Überspringe...")
                continue  # Überspringt zur nächsten E-Mail

            # Neuen absoluten Dateinamen erzeugen
            new_path_and_file_name = os.path.join(
                os.path.dirname(path_and_file_name),
                new_file_name
            )
            app_logger.debug(f"Neuer absoluter Dateiname: {new_path_and_file_name}.")

            # Jetzt msg-File umbenennen

            if not TEST_MODE:
                result = rename_file(path_and_file_name, new_path_and_file_name)
            else:
                result = FileOperationResult.TEST_MODE

            if result == FileOperationResult.SUCCESS:
                app_logger.debug(f"Die Umbenennung der Datei {path_and_file_name} war erfolgreich. Neuer Dateiname: {new_path_and_file_name}.")
            elif result == FileOperationResult.TEST_MODE:
                app_logger.debug(f"Keine Umbenennung der Datei {path_and_file_name} efolgt da Test-Modus. Neuer wäre Dateiname: {new_path_and_file_name}.")
            else:
                app_logger.error(f"Die Umbenennung der Datei {path_and_file_name} war nicht erfolgreich: {result}")
                continue  # Überspringt zur nächsten E-Mail

        return len(selected_emails)


    def save_as_msg(self, email: object, new_email_name=None, test_mode=False):
        """
        Speichert die ausgewählte E-Mail als MSG-Datei direkt über das Outlook-Objekt.

        :param email: `Email`-Objekt mit einem `outlook_item` (Original-Outlook-Objekt).
        """

        try:
            #outlook_item = email.outlook_item
            # Verbindung zu Outlook herstellen und auf den angegebenen Postfach-Namespace zugreifen
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            #outlook_item = namespace.GetItemFromID(email.entry_id, email.store_id)
            outlook_item = get_outlook_item_with_retry(namespace, email.entry_id, email.store_id)

            #if not outlook_item:
            if outlook_item is None:
                raise ValueError("Kein gültiges Outlook-Objekt gefunden.")
            else:
                pass

            # Verzeichnis normalisieren (entfernt gemischte Trennzeichen und leere Endungen)
            app_logger.debug(f"Verzeichnis aus 'self.export_directory': {self.export_directory}.")
            normalized_export_directory = os.path.normpath(self.export_directory)
            app_logger.debug(f"Verzeichnis nach Normalisierung: {normalized_export_directory}.")

            # Intialen Dateiname festlegen
            if not new_email_name:
                #subject = email.subject.replace(" ", "_")
                file_name = sanitize_filename(email.subject)
                sanitized_file_name = f"{file_name}.msg"
                app_logger.debug(f"Bereinigter Filename: {sanitized_file_name}.")

                file_path = os.path.join(normalized_export_directory, sanitized_file_name)
            else:
                file_path = os.path.join(normalized_export_directory, new_email_name + ".msg")

            # Speichern der Nachricht als MSG über Outlook
            if not test_mode:
                outlook_item.SaveAs(file_path, 3)  # 3 entspricht dem MSG-Format
                app_logger.debug(f"Datei gespeichert: {file_path}.")
            else:
                app_logger.warning(f"Datei nicht gespeichert, da Test-Modus aktiv.")

            return file_path

        except Exception as e:

            raise RuntimeError(f"MSG-Export fehlgeschlagen: {e}")



    # def save_as_pdf(self, email: object):
    #     """
    #     Speichert die ausgewählte E-Mail als PDF-Datei direkt über das Outlook-Objekt.
    #
    #     :param email: `Email`-Objekt mit einem `outlook_item` (Original-Outlook-Objekt).
    #     """
    #     if not self.export_path:
    #         raise RuntimeError("Kein Exportpfad festgelegt. Bitte rufe 'set_export_path()' auf.")
    #
    #     try:
    #         outlook_item = email.outlook_item
    #         if not outlook_item:
    #             raise ValueError("Kein gültiges Outlook-Objekt gefunden.")
    #
    #         # Dateiname festlegen
    #         subject = email.subject.replace(" ", "_")
    #         file_name = f"{subject}.pdf"
    #         file_path = os.path.join(self.export_path, file_name)
    #
    #         # Speichern der Nachricht über die Print-Methode von Outlook
    #         temp_html_path = os.path.join(self.export_path, f"{subject}.html")
    #
    #         # Nachricht als HTML speichern
    #         outlook_item.SaveAs(temp_html_path, 5)  # 5 entspricht dem HTML-Format
    #
    #         # HTML-Datei in PDF konvertieren (externes Tool oder Bibliothek erforderlich)
    #         self._convert_html_to_pdf(temp_html_path, file_path)
    #
    #         # Temporäre HTML-Datei löschen
    #         if os.path.exists(temp_html_path):
    #             os.remove(temp_html_path)
    #
    #         logging.info(f"E-Mail erfolgreich als PDF exportiert: {file_path}")
    #
    #     except Exception as e:
    #         logging.error(f"Fehler beim PDF-Export über Outlook: {e}")
    #         raise RuntimeError(f"PDF-Export fehlgeschlagen: {e}")
    #
    # def _convert_html_to_pdf(self, input_html: str, output_pdf: str):
    #     """
    #     Hilfsmethode zur Konvertierung von HTML zu PDF. Es wird eine externe Bibliothek verwendet.
    #
    #     :param input_html: Pfad zur HTML-Datei.
    #     :param output_pdf: Ziel-PDF-Datei.
    #     """
    #     try:
    #         from weasyprint import HTML  # Sicherstellen, dass `weasyprint` installiert ist.
    #         HTML(input_html).write_pdf(output_pdf)
    #         logging.info(f"HTML erfolgreich in PDF konvertiert: {output_pdf}")
    #     except ImportError as e:
    #         logging.error(f"Fehler: `weasyprint` ist nicht installiert: {e}")
    #         raise RuntimeError("HTML-zu-PDF-Konvertierung erfordert `weasyprint`.")
    #     except Exception as e:
    #         logging.error(f"HTML-zu-PDF-Konvertierung fehlgeschlagen: {e}")
    #         raise RuntimeError(f"HTML-zu-PDF-Konvertierung fehlgeschlagen: {e}")

def sanitize_filename(filename):
    """
    Bereinigt einen Dateinamen, indem unerlaubte Zeichen entfernt oder ersetzt werden.

    :param filename: Der ursprüngliche Dateiname (z. B. Subject einer E-Mail).
    :return: Bereinigter Dateiname, der sicher verwendet werden kann.
    """
    # Ersetze Leerzeichen mit Unterstrichen
    filename = filename.replace(" ", "_")

    # Entferne Zeichen, die in Dateinamen unzulässig sind
    filename = re.sub(r'[\\/*?:"<>|]', '_', filename)

    # Optional: Kürzen des Dateinamens, wenn er zu lang ist (z. B. max. 255 Zeichen)
    max_length = 255  # Maximale Dateinamenlänge (je nach Betriebssystem)
    if len(filename) > max_length:
        filename = filename[:max_length]

    return filename


def get_outlook_item_with_retry(namespace, entry_id, store_id, retries=3, delay=1):
    """
    Versucht, ein Outlook-Item anhand EntryID und StoreID mehrfach zu laden.
    Bei Fehlern werden mehrere Versuche unternommen, jeweils mit einer kurzen Pause.
    """
    for attempt in range(retries):
        try:
            return namespace.GetItemFromID(entry_id, store_id)
        except Exception as e:
            app_logger.warning(f"Versuch {attempt + 1} fehlgeschlagen: {e}")
            time.sleep(delay)
    app_logger.error(f"Outlook-Item konnte nach {retries} Versuchen nicht geladen werden (EntryID: {entry_id})")
    return None