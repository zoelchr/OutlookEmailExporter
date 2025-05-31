import os
import logging

class ExportManager:
    def __init__(self, email_table_model, export_directory):
        """
        Initialisiert den Exporter mit einem Tabellenmodell und einem Exportverzeichnis.

        :param email_table_model: Instanz von EmailTableModel.
        :param export_directory: Ausgewähltes Exportziel (z. B. Verzeichnis).
        """
        self.email_table_model = email_table_model
        self.export_directory = export_directory

        logging.info(f"ExportManager mit Tabellenmodell und Exportverzeichnis {export_directory} initialisiert.")


    def export_emails(self):
        """
        Exportiert die ausgewählten Emails in das Zielverzeichnis.
        """
        # 1. Abrufen der ausgewählten Emails
        selected_emails = self.email_table_model.get_selected_emails()

        if not selected_emails:
            print("Keine ausgewählten Emails zum Exportieren.")
            return

        # 2. Exportprozess - Hier können Sie die spezifische Exportimplementierung einfügen.
        print(f"Exportiere die folgenden Emails in das Verzeichnis: {self.export_directory}")
        for email in selected_emails:
            print(f"- {email}")  # Annahme: Email-Objekte haben eine sinnvoll definierte __str__-Methode

        # 3. Optional: Feedback an den Benutzer
        print("Export abgeschlossen.")


    # def save_as_msg(self, email: object):
    #     """
    #     Speichert die ausgewählte E-Mail als MSG-Datei direkt über das Outlook-Objekt.
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
    #         file_name = f"{subject}.msg"
    #         file_path = os.path.join(self.export_path, file_name)
    #
    #         # Speichern der Nachricht als MSG über Outlook
    #         outlook_item.SaveAs(file_path, 3)  # 3 entspricht dem MSG-Format
    #         logging.info(f"E-Mail erfolgreich als MSG exportiert: {file_path}")
    #
    #     except Exception as e:
    #         logging.error(f"Fehler beim MSG-Export über Outlook: {e}")
    #         raise RuntimeError(f"MSG-Export fehlgeschlagen: {e}")
    #
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