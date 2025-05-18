"""
outlook_connector.py

Dieses Modul stellt Funktionen zum Zugriff auf Outlook bereit (via COM-Interface `win32com`).
Aktuell wird eine Liste aller verfügbaren Postfächer im Profil zurückgegeben – unter Berücksichtigung
einer Ausschlussliste aus der Konfiguration (`IGNORE_POSTFAECHER`).

Fehler wie z. B. ein blockierter Outlook-Prozess oder ein fehlendes Profil werden robust behandelt
und im Log dokumentiert.
"""
import win32com.client
from logger import log
from config import IGNORE_POSTFAECHER
from config import EXCLUDE_FOLDERNAMES

def get_outlook_postfaecher():
    try:
        log("📥 Starte Zugriff auf Outlook", level=2)

        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as dispatch_error:
            log(f"❌ Outlook konnte nicht gestartet oder verbunden werden: {dispatch_error}", level=0)
            return []

        log("🔌 Outlook.Application instanziiert", level=2)

        try:
            namespace = outlook.GetNamespace("MAPI")
            stores = namespace.Stores
            log(f"📂 Anzahl Stores: {stores.Count}", level=2)
        except Exception as ns_error:
            log(f"❌ Fehler beim Zugriff auf Outlook-Namespace: {ns_error}", level=0)
            return []

        postfaecher = []

        for i in range(stores.Count):
            store = stores.Item(i + 1)
            log(f"🔍 Prüfe Store: {store.DisplayName}", level=3)
            if store.DisplayName not in IGNORE_POSTFAECHER:
                postfaecher.append(store.DisplayName)

        log(f"📬 {len(postfaecher)} Postfächer gefunden", level=2)
        return postfaecher

    except Exception as e:
        log(f"❌ Allgemeiner Fehler beim Outlook-Zugriff: {e}", level=0)
        return []


def get_outlook_ordner(postfach_name):
    """
    Ruft alle Outlook-Ordner eines Postfachs ab, die
    - mindestens eine E-Mail enthalten (Items.Count > 0),
    - keine ausgeschlossenen Namen enthalten (EXCLUDE_FOLDERNAMES).
    """
    try:
        log(f"📥 Beginne Ordnerabfrage für Postfach: {postfach_name}", level=2)

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders[postfach_name]

        ordnerliste = []

        def collect_folder_names(folder, path=""):
            full_path = f"{path}/{folder.Name}" if path else folder.Name
            folder_name_lower = folder.Name.lower()

            # Ausschluss durch Namensfilter
            if any(blocked in folder_name_lower for blocked in EXCLUDE_FOLDERNAMES):
                log(f"⛔️ Ordner ausgeschlossen (Name gefiltert): {full_path}", level=3)
                return

            # Nur Ordner mit Inhalt (E-Mails oder Elemente)
            try:
                item_count = folder.Items.Count
            except Exception:
                item_count = 0  # Bei Zugriffsfehlern ignorieren

            if item_count > 0:
                ordnerliste.append(full_path)
                log(f"✅ Ordner akzeptiert: {full_path} ({item_count} Elemente)", level=3)
            else:
                log(f"🚫 Ordner ohne Inhalt übersprungen: {full_path}", level=3)

            # Rekursiv prüfen
            for subfolder in folder.Folders:
                collect_folder_names(subfolder, full_path)

        collect_folder_names(root_folder)

        log(f"📂 {len(ordnerliste)} relevante Ordner für '{postfach_name}' gefunden", level=2)
        return ordnerliste

    except Exception as e:
        log(f"❌ Fehler beim Laden der Ordner für '{postfach_name}': {e}", level=0)
        return []
