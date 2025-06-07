"""
outlook_connector.py

Das Modul `outlook_connector.py` bietet eine zentrale Schnittstelle zur Interaktion mit Microsoft Outlook,
indem es das COM-Interface `win32com.client` verwendet. Es enthält grundlegende Funktionen für die Automatisierung
von Aufgaben bezogen auf Outlook, wie das Abrufen von Postfächern, das Durchsuchen von Ordnerstrukturen
und das Verwalten von E-Mail-Daten.

### Übersicht der Hauptfunktionen:
1. **Postfachverwaltung:**
   - Ermittelt alle verfügbaren Postfächer im aktuellen Outlook-Profil unter Berücksichtigung der Konfigurationsparameter,
     wie beispielsweise der Ausschlussliste (`IGNORE_POSTFAECHER`).
   - Sorgt für eine effiziente und gezielte Handhabung der Postfächer durch Filterung irrelevanter Elemente.

2. **Verzeichnis- und Ordnersuche:**
   - Ermöglicht die Rekursive Suche und Verarbeitung von Ordnerstrukturen in einem Postfach.
   - Respektiert dabei die maximale Suchtiefe (`MAX_FOLDER_DEPTH`) und eine Ausschlussliste (`EXCLUDE_FOLDERNAMES`) für Ordner, 
     um die Leistung zu optimieren und irrelevante Inhalte auszublenden.

3. **E-Mail-Verarbeitung:**
   - Bietet grundlegende Unterstützung für die Extraktion und strukturelle Handhabung von E-Mails in ausgewählten Ordnerstrukturen.

### Besondere Merkmale:
- **Robuste Fehlerbehandlung:** 
  - Tritt ein Problem mit der Outlook-Installation, einem blockierten Prozess oder einer fehlenden Konfiguration auf,
    werden diese Fehler detailliert protokolliert und gegebenenfalls an den Benutzer weitergeleitet, um eine stabile Ausführung
    sicherzustellen.
  
- **Flexibilität durch Konfiguration:**
  - Einstellungen wie ignorierte Postfächer (z. B. Archiv oder Testkonten), ausgenommene Ordnernamen oder die maximale Suchtiefe
    können zentral in der `.env`-Datei angegeben werden.

- **Integration in GUI-Systeme:**
  - Das Modul ist darauf ausgelegt, andere Systeme wie grafische Benutzeroberflächen zu unterstützen, wobei durch Log-Nachrichten und
    klare Schnittstellen eine einfache Integration möglich ist.
"""
import logging
app_logger = logging.getLogger(__name__)  
# Erstellt einen Logger für die aktuelle Datei. 
# Der Logger wird verwendet, um Debug- und Fehlerinformationen während der Ausführung zu protokollieren.

import win32com.client  
# Importiert die win32com.client-Bibliothek, die eine Schnittstelle für die Automatisierung von 
# Windows-Komponenten wie Outlook bereitstellt. Sie wird verwendet, um auf COM-Objekte zuzugreifen.

from config import IGNORE_POSTFAECHER, MAX_FOLDER_DEPTH, EXCLUDE_FOLDERNAMES  
# Importiert Konfigurationsparameter aus der `config.py`-Datei:
# - `IGNORE_POSTFAECHER`: Eine Liste von Postfächern, die nicht in der GUI angezeigt werden sollen.
# - `MAX_FOLDER_DEPTH`: Die maximale Tiefe, bis zu der in der Ordnerstruktur eines Postfachs gesucht wird.
# - `EXCLUDE_FOLDERNAMES`: Eine Liste von Ordnernamen, die aus der Suche ausgeschlossen werden.

from email_model import Email  
# Importiert die `Email`-Klasse aus dem Modul `email_model`, 


def get_outlook_postfaecher():
    """
    Ruft eine Liste der Postfächer ab, die im aktuellen Outlook-Profil verfügbar sind, 
    und filtert diese basierend auf einer Ignorierliste.

    Die Funktion verwendet die Microsoft Outlook COM-Schnittstelle, um auf Postfächer und gespeicherte Daten 
    zuzugreifen. Dabei werden alle Postfächer, die in der globalen Variable `IGNORE_POSTFAECHER` definiert sind, 
    aus der Rückgabe ausgeschlossen.

    Returns:
        list: Eine Liste der Namen von Postfächern, die für das aktuelle Profil zugänglich sind und 
              nicht in der Ignorierliste `IGNORE_POSTFAECHER` enthalten sind.

    Ablauf:
    1. Die Funktion stellt eine Verbindung zu den Outlook-Datenquellen her und ruft die verfügbaren 
       Speicherorte (sog. "Stores") ab.
    2. Für jeden Speicherort wird überprüft, ob der Name in der Ignorierliste `IGNORE_POSTFAECHER` enthalten ist.
       - Falls nicht, wird der Name des Speicherortes zu einer Liste hinzugefügt.
    3. Die gefilterte Liste der Postfächer wird zurückgegeben.

    Wichtige Hinweise:
    - Die Funktion ist von der Outlook COM-Schnittstelle abhängig und erfordert ein korrekt konfiguriertes 
      Outlook-Profil auf dem System, von dem der Zugriff erfolgt.
    """
    try:
        app_logger.debug("Starte Zugriff auf Outlook.")

        try:
            # Versuch, eine Instanz der Outlook-Anwendung mit der COM-Schnittstelle zu instanziieren.
            outlook = win32com.client.Dispatch("Outlook.Application")
        except Exception as dispatch_error:
            # Fehlerprotokollierung, falls Outlook nicht gestartet oder verbunden werden kann.
            app_logger.error(f"Outlook konnte nicht gestartet oder verbunden werden: {dispatch_error}")
            return []  # Rückgabe einer leeren Liste, da keine Aktion möglich ist.

        # Protokollierung, dass die Instanz der Outlook-Anwendung erfolgreich erstellt wurde.
        app_logger.debug("Outlook.Application instanziiert.")

        try:
            # Zugriff auf das MAPI-Namespace-Objekt von Outlook, das die Schnittstelle für Postfächer, Ordner
            # und Elemente innerhalb des aktuellen Benutzerprofils bereitstellt.
            namespace = outlook.GetNamespace("MAPI")
            
            # Abrufen der Sammlungen aller verfügbaren "Stores" (Postfächer und Speicherorte), die mit dem aktuellen
            # Outlook-Profil verknüpft sind (z. B. Exchange-Postfächer, IMAP-Ordner, Archivdateien).
            stores = namespace.Stores

            # Protokollierung der Anzahl der gefundenen Postfächer/Speicherorte.
            app_logger.debug(f"Anzahl gefundener Outlook-Stores: {stores.Count}")

            # Protokollierung der Namen der gefundenen Stores zur Nachverfolgbarkeit, um sicherzustellen,
            # dass die Verbindung korrekt arbeitet und alle relevanten Stores erkannt wurden.
            app_logger.debug(f"Gefundene Outlook-Stores: {[store.DisplayName for store in stores]}")

        except Exception as ns_error:
            # Fehlerprotokollierung, falls beim Zugriff auf das MAPI-Namespaces-Objekt ein Problem auftritt.
            app_logger.error(f"Fehler beim Zugriff auf Outlook-Namespace: {ns_error}")
            return []  # Rückgabe einer leeren Liste, da keine Aktion möglich ist.

        # Initialisiere eine leere Liste zur Speicherung der gefundenen Postfächer
        postfaecher = []

        # Durchlaufe die Sammlung aller Stores (Postfächer/Speicher), die aus Outlook geladen wurden
        for i in range(stores.Count):
            # Greift auf das aktuelle Store-Element zu (1-basiert, daher `i + 1`)
            store = stores.Item(i + 1)
            # Protokolliere den Namen des aktuellen Stores, der geprüft wird
            app_logger.debug(f"Prüfe Store: {store.DisplayName}")
            
            # Überprüfe, ob der Name des Stores nicht in der Liste der zu ignorierenden Postfächer (IGNORE_POSTFAECHER) enthalten ist
            if store.DisplayName not in IGNORE_POSTFAECHER:
                # Wenn der Store nicht in der Ignorier-Liste ist, füge ihn der Postfachliste hinzu
                postfaecher.append(store.DisplayName)

        # Protokolliere die Anzahl der gefundenen Postfächer, die nicht ignoriert wurden
        app_logger.debug(f"{len(postfaecher)} Postfächer gefunden")

        # Gib die gefilterte Liste der Postfächer zurück
        return postfaecher

    except Exception as e:
        app_logger.error(f"Allgemeiner Fehler beim Outlook-Zugriff: {e}")
        return []


def get_outlook_ordner(postfach_name):
    """
    Ruft eine Liste aller relevanten Ordner eines angegebenen Outlook-Postfachs ab.

    Ablauf:
    1. Initialisiert den Zugriff auf die Outlook-Anwendung und das angegebene Postfach.
    2. Prüft die Ordner des Postfachs:
        - Es werden nur Ordner berücksichtigt, die mindestens eine E-Mail enthalten (`Items.Count > 0`).
        - Ordner, deren Namen in `EXCLUDE_FOLDERNAMES` definiert sind, werden ausgeschlossen.
    3. Ruft alle gültigen Ordnernamen rekursiv bis zu einer maximalen Tiefe (`MAX_FOLDER_DEPTH`) ab.
    4. Gibt die Liste der relevanten Ordner zurück.

    Parameter:
        postfach_name (str): Der Name des Outlook-Postfachs, dessen Ordner analysiert werden sollen.

    Returns:
        list: Eine Liste von Ordnernamen, die den angegebenen Kriterien entsprechen. 
              Gibt eine leere Liste zurück, falls ein Fehler auftritt.

    Fehlerbehandlung:
    - Fängt alle Ausnahmen während des Zugriffs auf die Outlook-APIs ab.
    - Gibt bei Fehlern eine leere Liste zurück und loggt eine entsprechende Fehlermeldung.
    """
    try:
        # Protokolliert den Start der Ordnerabfrage
        app_logger.debug(f"Beginne Ordnerabfrage für Postfach: {postfach_name}")

        # Erstellt eine Verbindung zur Outlook-Anwendung
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")

        # Ruft den Wurzelordner des angegebenen Postfachs ab
        root_folder = namespace.Folders[postfach_name]

        # Initialisiert eine leere Liste für die relevanten Ordner
        ordnerliste = []

        # Sammelt rekursiv alle relevanten Ordnernamen (Filter- und Tiefenkontrolle inkludiert)
        ordnerliste = collect_folder_names(
            root_folder, 
            exclude=EXCLUDE_FOLDERNAMES, 
            max_depth=MAX_FOLDER_DEPTH
        )

        # Protokolliert die Anzahl gefundener relevanter Ordner
        app_logger.info(f"{len(ordnerliste)} relevante Ordner für '{postfach_name}' gefunden")

        # Gibt die Liste der relevanten Ordner zurück
        return ordnerliste

    except Exception as e:
        # Protokolliert Fehler bei der Ordnerabfrage
        app_logger.error(f"Fehler beim Laden der Ordner für '{postfach_name}': {e}")
        # Gibt eine leere Liste bei einem Fehler zurück
        return []


def collect_folder_names(folder, path="", exclude=None, max_depth=float('inf'), current_depth=0):
    """
    Rekursive Funktion, um Ordnernamen ab einer gegebenen Ordnerstruktur zu sammeln.
    Berücksichtigt optional eine maximale Ordnertiefe.

    Args:
        folder: Der aktuelle Ordner, dessen Subordner durchsucht werden sollen.
        path: Der aktuelle Pfad (für vollständige Pfadangaben im Ergebnis).
        exclude: Eine Liste von zu ignorierenden Ordnernamen (optional).
        max_depth: Die maximale Ordnertiefe (optional, Standard: kein Limit).
        current_depth: Die aktuelle Tiefe der Rekursion (intern verwendet, Standard: 0).

    Returns:
        list: Eine Liste aller gefundenen Ordnerpfade.
    """
    folder_list = []
    full_path = f"{path}/{folder.Name}" if path else folder.Name
    folder_name_lower = folder.Name.lower()

    # Ausschluss durch Namensfilter
    if exclude and any(blocked in folder_name_lower for blocked in exclude):
        app_logger.debug(f"Ordner ausgeschlossen (Name gefiltert): {full_path}")
        return folder_list

    # Überprüfe die maximale Tiefe
    if current_depth > max_depth:
        app_logger.debug(f"Maximale Tiefe erreicht ({current_depth}/{max_depth}) für Ordner: {full_path}")
        return folder_list

    # Nur Ordner mit Inhalt
    try:
        item_count = folder.Items.Count
    except Exception:
        item_count = 0  # Zugriffsfehler ignorieren

    if item_count > 0:
        folder_list.append(full_path)
        app_logger.debug(f"Ordner akzeptiert: {full_path} ({item_count} Elemente)")
    else:
        app_logger.debug(f"Ordner ohne Inhalt übersprungen: {full_path}")

    # Rekursion: Gehe nur tiefer, wenn die maximale Tiefe noch nicht erreicht ist
    if current_depth < max_depth:
        for subfolder in folder.Folders:
            folder_list.extend(collect_folder_names(
                subfolder,
                full_path,
                exclude,
                max_depth,
                current_depth + 1
            ))

    return folder_list


def lade_emails(postfach_name: str, ordner_pfad: str) -> list[Email]:
    """
    Lädt E-Mails aus einem angegebenen Outlook-Postfach und Ordner mithilfe des COM-Interfaces.

    Args:
        postfach_name (str): Der Anzeigename des Postfachs, das ausgewählt wurde (wie in der GUI-ComboBox dargestellt).
        ordner_pfad (str): Pfad zum Zielordner innerhalb des Postfachs, z. B. "Posteingang/Projekte/2025".

    Returns:
        list[Email]: Eine Liste von `Email`-Objekten, die die wichtigsten E-Mail-Daten enthalten (z. B. Betreff, Absender, Empfangszeit).
                     Die `Email`-Objekte entsprechen der Klasse `Email` aus `email_model.py`.
                     Gibt eine leere Liste zurück, falls keine E-Mails gefunden werden oder ein Fehler auftritt.
    """
    emails = []  # Liste zum Speichern der extrahierten E-Mails

    try:
        # Logs einen Hinweis, dass der Ladevorgang für ein spezifisches Postfach und einen Ordner beginnt.
        app_logger.info(f"Lese Mails aus: Postfach='{postfach_name}' | Ordner='{ordner_pfad}'")

        # Verbindung zu Outlook herstellen und auf den angegebenen Postfach-Namespace zugreifen
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders[postfach_name]

        # Navigiere durch die Ordnerstruktur, beginnend beim root_folder,
        # und aktualisiere das aktuelle Zielverzeichnis anhand der Segmente des Ordnerpfades.
        # Dabei wird nur der gewählte Ordner verarbeitet, Unterordner werden nicht berücksichtigt.
        current_folder = root_folder
        for part in ordner_pfad.split("/"):
            current_folder = current_folder.Folders[part]

        # Sortiere die E-Mails im Ordner nach Empfangszeit absteigend
        #current_folder.Items.Sort("[ReceivedTime]", True)

        app_logger.debug(f"Ordner: {current_folder.Name}, DefaultItemType: {getattr(current_folder, 'DefaultItemType', 'unbekannt')}")

        # Iteriere über die Objekte im Ordner und extrahiere nur E-Mail-Objekte (Class = 43)
        for item in current_folder.Items:
            try:
                if hasattr(item, "Class") and item.Class == 43:  # 43 steht für E-Mail-Objekt (olMail)
                    app_logger.debug(f"Objekt hat das Class-Attribute 43 (E-Mail)")
                    email = Email(
                        received=str(item.ReceivedTime.strftime("%d.%m.%Y %H:%M")),     # Empfangszeit als formatierter String
                        sender_name=item.SenderName,                                    # Name des Absenders
                        sender_email=item.SenderEmailAddress,                           # E-Mail-Adresse des Absenders
                        subject=item.Subject or "",                                     # Betreff (oder leer, falls nicht vorhanden)
                        #outlook_item=item,                                               # Das Original-Outlook-Objekt
                        entry_id = item.EntryID,
                        store_id = item.Parent.StoreID
                    )
                    emails.append(email)  # Füge das generierte `Email`-Objekt der Liste hinzu
            except Exception as e:
                app_logger.debug(f"Item konnte nicht geprüft werden: {e}")

        # Logs die Anzahl der erfolgreich geladenen E-Mails
        app_logger.info(f"{len(emails)} Mails geladen.")
        return emails  # Gibt die gefüllte Liste zurück, falls erfolgreich

    except Exception as e:
        # Logs den Fehler, falls während des Ladevorgangs ein Problem auftritt
        app_logger.error(f"Fehler beim Lesen von Mails: {e}")
        return []  # Gibt eine leere Liste zurück, falls ein Fehler auftritt


def count_emails_in_folder(postfach_name: str, ordner_pfad: str) -> int:
    """
    Gibt die Anzahl der E-Mails (olMailItem, Class == 43) im angegebenen Outlook-Postfach und Ordner zurück.
    Bei Fehlern wird -1 zurückgegeben.

    Args:
        postfach_name (str): Name des Outlook-Postfachs.
        ordner_pfad (str): Pfad zum gewünschten Ordner (z.B. "Posteingang" oder "Unterordner1/Unterordner2").

    Returns:
        int: Anzahl der E-Mails im Ordner, oder -1 bei Fehler.
    """
    try:
        import win32com.client

        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        root_folder = namespace.Folders[postfach_name]

        # Navigiere durch die Ordnerstruktur anhand des Pfades
        current_folder = root_folder
        for part in ordner_pfad.split("/"):
            if part:
                current_folder = current_folder.Folders[part]

        # Zähle nur Items mit Class == 43 (olMailItem)
        mail_count = sum(1 for item in current_folder.Items if getattr(item, "Class", None) == 43)
        return mail_count

    except Exception as e:
        # Optional: Logging hier ergänzen, z.B. mit app_logger.error(f"Fehler beim Zählen der E-Mails: {e}")
        return -1