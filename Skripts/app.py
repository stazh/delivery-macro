import sys
import os
import importlib
import argparse
import tkinter as tk
from tkinter import messagebox
from threading import Thread

# Globales Root-Fenster für alle MessageBoxen
root = tk.Tk()
root.withdraw()
root.attributes("-topmost", True)

# Aktuelles Arbeitsverzeichnis
current_dir = os.path.dirname(os.path.abspath(__file__))

# Pfade für Bibliotheken und Dienste
services_path = os.path.join(current_dir, 'services')
config_path = os.path.join(current_dir.replace('\main.pyz', ''))

import config
if getattr(sys, 'frozen', False):
    scripts_dir = os.path.dirname(sys.executable).replace(config.SKRIPTS_FOLDER_PATH, '')
    PROJECT_DIR = os.path.dirname(scripts_dir)
else:
    scripts_dir = os.path.dirname(os.path.abspath(__file__)).replace(config.SKRIPTS_FOLDER_PATH, '')
    PROJECT_DIR = os.path.dirname(scripts_dir)

os.chdir(PROJECT_DIR)

# Kommando-Mapping für verschiedene Aktionen
COMMANDS = {
    'import-file-offer': ('fileOfferService', 'import_file_offer_word', 'Importiert Aktenangebot (Word Datei) in die Excel-Tabelle (Angebot_aktuell)'),
    'create-template': ('fileOfferService', 'create_template', 'Erstellt eine Excel-Vorlage (Angebot_aktuell)'),
    'create-delivery-agreement': ('deliveryAgreementService', 'create_delivery_agreement', 'Erstellt eine Word-Datei für die Ablieferungsvereinbarung'),
    'import-delivery-list-word': ('deliveryListService', 'import_delivery_list_word', 'Importiert die Ablieferungsvereinbarung und erstellt eine Excel-Datei für das Ablieferungsverzeichnis'),
    'import-delivery-list-excel': ('deliveryListService', 'import_delivery_list_excel', 'Importiert die Daten aus Angebot_aktuell und erstellt eine Excel-Datei für das Ablieferungsverzeichnis'),
}

def run_with_loading(command: str):
    """Zeigt ein Ladefenster an, während die angegebene Funktion im Hintergrund ausgeführt wird."""
    rootLoading = tk.Toplevel(root)
    rootLoading.title("Loading")
    rootLoading.geometry("250x100")
    rootLoading.attributes("-topmost", True)
    rootLoading.resizable(False, False)
    tk.Label(rootLoading, text="Bitte warten...").pack(expand=True)

    root.update_idletasks()
    rootLoading.update_idletasks()
    rootLoading.deiconify()

    task_done = {'finished': False}

    def task():
        execute_command(command)
        task_done['finished'] = True

    Thread(target=task, daemon=True).start()

    def check_task():
        if task_done['finished']:
            rootLoading.destroy()
        else:
            rootLoading.after(100, check_task)

    check_task()
    rootLoading.mainloop()

def execute_command(command: str) -> None:
    """Führt das angegebene Kommando aus, indem die zugehörige Funktion importiert und aufgerufen wird."""
    if command in COMMANDS:
        file_name, function_name, _ = COMMANDS[command]
        try:
            module = importlib.import_module(file_name)  # Modul dynamisch importieren
            func = getattr(module, function_name)  # Funktion aus dem Modul holen
            func()  # Funktion ausführen
        except ModuleNotFoundError as e:
            messagebox.showerror("Fehler", f"Die Datei '{file_name}.py' wurde im services-Ordner nicht gefunden. Fehler: {e}", parent=root)
        except AttributeError as e:
            messagebox.showerror("Fehler", f"Die Funktion '{function_name}' wurde in der Datei '{file_name}.py' nicht gefunden. Fehler: {e}", parent=root)
        except Exception as e:
            messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten: {e}", parent=root)
    else:
        messagebox.showerror("Fehler", f"Der Befehl '{command}' ist nicht im Mapping definiert.", parent=root)

def main() -> None:
    """Hauptfunktion, die das Kommando aus der Kommandozeile verarbeitet und ausführt."""
    try:
        # Pfade zu den Bibliotheken und Diensten zum sys.path hinzufügen
        sys.path.append(services_path)
        sys.path.append(config_path)

        # Kommandozeilenparser initialisieren
        parser = argparse.ArgumentParser(description="CLI für das Ablieferungsmakro: Ruft die passende Funktion der richtigen Datei auf.")
        parser.add_argument('command', type=str, help='Der Befehl, der die zugehörige Datei und Funktion bestimmt.')
        args = parser.parse_args()
        
        # Funktion basierend auf dem Kommando ausführen
        run_with_loading(args.command)
        root.destroy()
    except Exception as e:
        root.destroy()
        messagebox.showerror("Fehler", f'Es ist ein Fehler aufgetreten: {e}', parent=root)

# Startet das Skript, wenn es direkt ausgeführt wird
if __name__ == '__main__':
    main()