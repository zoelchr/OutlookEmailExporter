"""
email_model.py

Dieses Modul definiert die Datenstruktur für eine E-Mail, wie sie im Outlook-Exporttool
verwendet wird. Die Klasse `Email` bildet die relevante Darstellung für die GUI und 
für Exportfunktionen ab.
"""
from dataclasses import dataclass

@dataclass
class Email:
    received: str         # Datum und Uhrzeit als formatierter String
    sender_name: str      # Anzeigename des Absenders
    sender_email: str     # E-Mail-Adresse des Absenders
    subject: str          # Betreff der Nachricht
    #outlook_item: object  # Original Outlook-Objekt (MailItem) für spätere Aktionen (Export etc.)
    entry_id: str
    store_id: str
