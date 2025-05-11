import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import win32com.client
import pythoncom
import os

global_messages = []
global_ordner_mapping = {}

def get_outlook_folders():
    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch("Outlook.Application")
    namespace = outlook.GetNamespace("MAPI")
    folders = [namespace.Folders.Item(i + 1).Name for i in range(namespace.Folders.Count)]
    pythoncom.CoUninitialize()
    return folders

def get_ordner_im_postfach(postfachname):
    global global_ordner_mapping
    pythoncom.CoInitialize()
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        namespace = outlook.GetNamespace("MAPI")
        for i in range(namespace.Folders.Count):
            postfach = namespace.Folders.Item(i + 1)
            if postfach.Name == postfachname:
                ordnernamen = []
                global_ordner_mapping = {}  # Clear old mappings

                for folder in postfach.Folders:
                    ordnernamen.append(folder.Name)
                    global_ordner_mapping[folder.Name] = folder
                    for subfolder in folder.Folders:
                        name = f"{folder.Name}/{subfolder.Name}"  # Pfad darstellen
                        ordnernamen.append(name)
                        global_ordner_mapping[name] = subfolder
                pythoncom.CoUninitialize()
                return ordnernamen
    except Exception as e:
        messagebox.showerror("Fehler", f"Ordner konnten nicht geladen werden:\n{e}")
        pythoncom.CoUninitialize()
        return []
    return []

def get_mails(folder):
    pythoncom.CoInitialize()
    try:
        messages = folder.Items
        messages.Sort("[ReceivedTime]", True)
        mails = []
        for i, msg in enumerate(messages):
            if i >= 100:
                break
            betreff = msg.Subject or "(Ohne Betreff)"
            sender = msg.SenderName or "(Unbekannt)"
            datum = msg.SentOn.strftime("%d.%m.%Y %H:%M") if msg.SentOn else "(Kein Datum)"
            mails.append((betreff, sender, datum, msg))
        pythoncom.CoUninitialize()
        return mails
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Zugriff auf Mails:\n{e}")
        pythoncom.CoUninitialize()
        return []
    return []

def anzeigen_mails():
    global global_messages
    for row in tree.get_children():
        tree.delete(row)

    ordnerwahl = combo_ordner.get()
    if ordnerwahl not in global_ordner_mapping:
        messagebox.showwarning("Ungültiger Ordner", "Bitte gültigen Ordner wählen.")
        return

    folder = global_ordner_mapping[ordnerwahl]
    global_messages = get_mails(folder)

    if not global_messages:
        label_info.config(text="Keine Mails gefunden.")
        return

    label_info.config(text=f"{len(global_messages)} Mails gefunden:")
    for i, (betreff, sender, datum, _) in enumerate(global_messages):
        tree.insert("", "end", iid=str(i), values=(sender, datum, betreff))

def ausgewaehlte_speichern():
    if not global_messages:
        messagebox.showinfo("Keine Mails", "Keine Mails geladen.")
        return

    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Keine Auswahl", "Bitte mindestens eine E-Mail auswählen.")
        return

    ordner = filedialog.askdirectory()
    if not ordner:
        return

    gespeichert = 0
    for item_id in selected_items:
        i = int(item_id)
        betreff, _, _, msg = global_messages[i]
        dateiname = "".join(c for c in betreff if c not in r'\/:*?"<>|').strip()
        pfad = os.path.join(ordner, f"{dateiname[:80] or 'Mail'}_{i+1}.msg")
        try:
            msg.SaveAs(pfad)
            gespeichert += 1
        except Exception as e:
            print(f"Fehler beim Speichern von Mail {i+1}: {e}")

    messagebox.showinfo("Fertig", f"{gespeichert} E-Mails gespeichert.")

def lade_ordner(event=None):
    postfachname = combo_postfach.get()
    if not postfachname:
        return
    ordnernamen = get_ordner_im_postfach(postfachname)
    combo_ordner["values"] = ordnernamen
    if ordnernamen:
        combo_ordner.current(0)

# GUI
fenster = tk.Tk()
fenster.title("Outlook Mail-Manager – Dynamische Ordner")
fenster.geometry("950x650")

# Dropdown für Postfach
tk.Label(fenster, text="Postfach auswählen:").pack(pady=5)
combo_postfach = ttk.Combobox(fenster, width=50, values=get_outlook_folders())
combo_postfach.pack()
combo_postfach.bind("<<ComboboxSelected>>", lade_ordner)

# Dropdown für Ordner
tk.Label(fenster, text="Ordner auswählen:").pack(pady=5)
combo_ordner = ttk.Combobox(fenster, width=70)
combo_ordner.pack()

# Button zum Anzeigen
tk.Button(fenster, text="Mails anzeigen", command=anzeigen_mails).pack(pady=10)

# Info-Label
label_info = tk.Label(fenster, text="")
label_info.pack()

# Treeview (Tabelle)
columns = ("Absender", "Versanddatum", "Betreff")
tree = ttk.Treeview(fenster, columns=columns, show="headings", selectmode="extended")
for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=250 if col != "Versanddatum" else 120)

tree.pack(fill="both", expand=True, padx=10, pady=10)

# Button zum Speichern
tk.Button(fenster, text="Ausgewählte speichern", command=ausgewaehlte_speichern).pack(pady=10)

fenster.mainloop()
