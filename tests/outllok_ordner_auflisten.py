import win32com.client
import pythoncom

def ordner_auflisten(postfach_name):
    pythoncom.CoInitialize()

    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    postfach_name = "ruediger@zoelch.me"

    for i in range(namespace.Folders.Count):
        postfach = namespace.Folders.Item(i + 1)
        if postfach.Name == postfach_name:
            print(f"📂 Postfach gefunden: {postfach.Name}")
            # Hauptordner auflisten
            for folder in postfach.Folders:
                print(f"- {folder.Name}")
                # Falls es Unterordner gibt, diese auch auflisten
                for subfolder in folder.Folders:
                    print(f"  -> {subfolder.Name}")
            break
    else:
        print(f"❌ Postfach '{postfach_name}' nicht gefunden.")

    pythoncom.CoUninitialize()

# Beispielaufruf
postfach_name = "Rüdiger Zölch"  # <- Hier deinen Postfachnamen anpassen
ordner_auflisten(postfach_name)