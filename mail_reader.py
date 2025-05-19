import logging
app_logger = logging.getLogger(__name__)
"""
mail_reader.py

Dieses Modul extrahiert E-Mails aus einem ausgewählten Outlook-Ordner und wandelt sie
in `Email`-Objekte um (siehe `email_model.py`).

Die Mails können dann z. B. in einer Tabelle angezeigt, ausgewählt oder exportiert werden.
"""

import win32com.client
from email_model import Email

def lade_emails(postfach_name: str, ordner_pfad: str) -> list[Email]:
    """
    Lädt E-Mails aus einem angegebenen Outlook-Postfach und Ordner.

    Args:
        postfach_name (str): Der Anzeigename des Postfachs (wie in der ComboBox).
        ordner_pfad (str): Pfad zum Ordner, z. B. "Posteingang/Projekte/2025"

    Returns:
        list[Email]: Liste von E-Mail-Objekten für die Anzeige oder Weiterverarbeitung
    """
    emails = []

    try:
        app_logger.info(f"📨 Lese Mails aus: Postfach='{postfach_name}' | Ordner='{ordner_pfad}'")

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders[postfach_name]

        # Ordnerpfad rekursiv auflösen
        current_folder = root_folder
        for part in ordner_pfad.split("/"):
            current_folder = current_folder.Folders[part]

        # Nur Mail-Objekte (optional: Items.Sort("ReceivedTime", True) für neueste oben)
        for item in current_folder.Items:
            if item.Class == 43:  # 43 = olMail
                email = Email(
                    received=str(item.ReceivedTime.strftime("%d.%m.%Y %H:%M")),
                    sender_name=item.SenderName,
                    sender_email=item.SenderEmailAddress,
                    subject=item.Subject or "",
                    outlook_item=item
                )
                emails.append(email)

        app_logger.info(f"📧 {len(emails)} Mails geladen")
        return emails

    except Exception as e:
        app_logger.error(f"❌ Fehler beim Lesen von Mails: {e}")
        return []
