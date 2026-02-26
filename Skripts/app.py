import sys
import os
import importlib
import argparse
from tkinter import messagebox

currentDir = os.path.dirname(os.path.abspath(__file__))

libs_path = os.path.join(currentDir, 'libs')
services_path = os.path.join(currentDir, 'services')
config_path = os.path.join(currentDir.replace('\main.pyz', ''))

COMMANDS = {
    'import-file-offer': ('fileOfferService', 'importFileOfferWord', 'Importiert Aktenangebot (Word Datei) in die Excel tabelle (Angebot_aktuell)'),
    'create-template': ('fileOfferService', 'createTemplate', 'Erstellt eine Excel Vorlage (Angebot_aktuell)'),
    'create-delivery-agreement': ('deliveryAgreementService', 'createDeliveryAgreement', 'Erstellt eine Word Datei für die Ablieferungsvereinbarung'),
    'import-delivery-list-word': ('deliveryListService', 'importDeliveryListWord', 'Importiert die Ablieferungsvereinbarung und erstellt eine Excel Datei für das Ablieferungsverzeichnis'),
    'import-delivery-list-excel': ('deliveryListService', 'importDeliveryListExcel', 'Importiert die Daten aus Angebot_aktuell und erstell eine Excel Datei für das Ablieferungsverzeichnis'),
}

def callFunction(command):
    if command in COMMANDS:
        fileName, functionName, _ = COMMANDS[command]
        try:
            module = importlib.import_module(fileName)
            func = getattr(module, functionName)
            func()
        except ModuleNotFoundError as e:
            messagebox.showerror("Fehler", f"Die Datei '{fileName}.py' wurde im services-Ordner nicht gefunden. Error: {e}")
        except AttributeError as e:
            messagebox.showerror("Fehler", f"Die Funktion '{functionName}' wurde in der Datei '{fileName}.py' nicht gefunden. Error: {e}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten: {e}")
    else:
        messagebox.showerror("Fehler", f"Der Befehl '{command}' ist nicht im Mapping definiert.")

def main():
    try:
        sys.path.append(services_path)
        sys.path.append(libs_path)
        sys.path.append(config_path)
        parser = argparse.ArgumentParser(description="CLI ist für das Ablieferungsmakro und ruft die richtige Funktion der richtigen Datei auf.")
        parser.add_argument('command', type=str, help='Der Befehl, der die zugehörige Datei und Funktion bestimmt.')
        args = parser.parse_args()
        callFunction(args.command)
    except Exception as e:
        messagebox.showerror("Fehler", f'{e}')

if __name__ == '__main__':
    main()