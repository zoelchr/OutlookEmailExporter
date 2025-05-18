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
