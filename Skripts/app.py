import sys
import os
import importlib
import argparse
from tkinter import messagebox

services_path = os.path.join(os.path.dirname(__file__), 'services')
sys.path.append(services_path)

COMMANDS = {
    'import-file-offer': ('fileOfferService', 'importFileOfferWord', 'Importiert Aktenangebot (Word Datei) in die Excel tabelle (Angebot_aktuell)'),
    'create-template': ('fileOfferService', 'createTemplate', 'Erstellt eine Excel Vorlage (Angebot_aktuell)'),
    'create-delivery-agreement': ('deliveryAgreementService', 'createDeliveryAgreement', 'Erstellt eine Word Datei für die Ablieferungsvereinbarung'),
    'import-delivery-list-word': ('deliveryListService', 'importDeliveryListWord', 'Importiert die Ablieferungsvereinbarung und erstellt eine Excel Datei für das Ablieferungsverzeichnis'),
    'import-delivery-list-excel': ('deliveryListService', 'importDeliveryListExcel', 'Importiert die Daten aus Angebot_aktuell und erstell eine Excel Datei für die Ablieferungsverzeichnis'),
}


def callFunction(command):
    if command in COMMANDS:
        fileName, functionName, _ = COMMANDS[command]

        try:
            print(f"Versuche, {fileName}.py zu importieren...")
            module = importlib.import_module(f'services.{fileName}')
            func = getattr(module, functionName)
            func()
        except ModuleNotFoundError as e:
            print(f"Fehler: Die Datei '{fileName}.py' wurde im services-Ordner nicht gefunden.")
            print(f"Details: {e}")
        except AttributeError as e:
            print(f"Fehler: Die Funktion '{functionName}' wurde in der Datei '{fileName}.py' nicht gefunden.")
            print(f"Details: {e}")
        except Exception as e:
            print(f"Es ist ein Fehler aufgetreten: {e}")
    else:
        print(f"Fehler: Der Befehl '{command}' ist nicht im Mapping definiert.")

def main():
    messagebox.showinfo("Main aufgerufen", "Dies ist eine einfache Nachricht.")
    parser = argparse.ArgumentParser(description="CLI ist für das Ablieferungsmakro und ruft die richtige Funktion der richtigen Datei auf.")
    parser.add_argument('command', type=str, help='Der Befehl, der die zugehörige Datei und Funktion bestimmt.')
    args = parser.parse_args()
    callFunction(args.command)

if __name__ == '__main__':
    main()
