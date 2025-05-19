import logging
app_logger = logging.getLogger(__name__)
"""
outlook_connector.py

Dieses Modul stellt Funktionen zum Zugriff auf Outlook bereit (via COM-Interface `win32com`).
Aktuell wird eine Liste aller verfügbaren Postfächer im Profil zurückgegeben – unter Berücksichtigung
einer Ausschlussliste aus der Konfiguration (`IGNORE_POSTFAECHER`).

Fehler wie z. B. ein blockierter Outlook-Prozess oder ein fehlendes Profil werden robust behandelt
und im Log dokumentiert.
"""
import win32com.client
from config import IGNORE_POSTFAECHER
from config import EXCLUDE_FOLDERNAMES

def get_outlook_postfaecher():
    try:
        app_logger.info("📥 Starte Zugriff auf Outlook")

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as dispatch_error:
            app_logger.error(f"❌ Outlook konnte nicht gestartet oder verbunden werden: {dispatch_error}")
            return []

        app_logger.info("🔌 Outlook.Application instanziiert")

        try:
            namespace = outlook.GetNamespace("MAPI")
            stores = namespace.Stores
            app_logger.info(f"📂 Anzahl Stores: {stores.Count}")
        except Exception as ns_error:
            app_logger.error(f"❌ Fehler beim Zugriff auf Outlook-Namespace: {ns_error}")
            return []

        postfaecher = []

        for i in range(stores.Count):
            store = stores.Item(i + 1)
            app_logger.debug(f"🔍 Prüfe Store: {store.DisplayName}")
            if store.DisplayName not in IGNORE_POSTFAECHER:
                postfaecher.append(store.DisplayName)

        app_logger.info(f"📬 {len(postfaecher)} Postfächer gefunden")
        return postfaecher

    except Exception as e:
        app_logger.error(f"❌ Allgemeiner Fehler beim Outlook-Zugriff: {e}")
        return []


def get_outlook_ordner(postfach_name):
    """
    Ruft alle Outlook-Ordner eines Postfachs ab, die
    - mindestens eine E-Mail enthalten (Items.Count > 0),
    - keine ausgeschlossenen Namen enthalten (EXCLUDE_FOLDERNAMES).
    """
    try:
        app_logger.info(f"📥 Beginne Ordnerabfrage für Postfach: {postfach_name}")

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders[postfach_name]

        ordnerliste = []

        def collect_folder_names(folder, path=""):
            full_path = f"{path}/{folder.Name}" if path else folder.Name
            folder_name_lower = folder.Name.lower()

            # Ausschluss durch Namensfilter
            if any(blocked in folder_name_lower for blocked in EXCLUDE_FOLDERNAMES):
                app_logger.debug(f"⛔️ Ordner ausgeschlossen (Name gefiltert): {full_path}")
                return

            # Nur Ordner mit Inhalt (E-Mails oder Elemente)
            try:
                item_count = folder.Items.Count
            except Exception:
                item_count = 0  # Bei Zugriffsfehlern ignorieren

            if item_count > 0:
                ordnerliste.append(full_path)
                app_logger.debug(f"✅ Ordner akzeptiert: {full_path} ({item_count} Elemente)")
            else:
                app_logger.debug(f"🚫 Ordner ohne Inhalt übersprungen: {full_path}")

            # Rekursiv prüfen
            for subfolder in folder.Folders:
                collect_folder_names(subfolder, full_path)

        collect_folder_names(root_folder)

        app_logger.info(f"📂 {len(ordnerliste)} relevante Ordner für '{postfach_name}' gefunden")
        return ordnerliste

    except Exception as e:
        app_logger.error(f"❌ Fehler beim Laden der Ordner für '{postfach_name}': {e}")
        return []
