import sys
import os
import importlib
import argparse
from tkinter import messagebox

# Aktuelles Arbeitsverzeichnis
current_dir = os.path.dirname(os.path.abspath(__file__))

# Pfade für Bibliotheken und Dienste
services_path = os.path.join(current_dir, 'services')
config_path = os.path.join(current_dir.replace('\main.pyz', ''))

# Kommando-Mapping für verschiedene Aktionen
COMMANDS = {
    'import-file-offer': ('fileOfferService', 'import_file_offer_word', 'Importiert Aktenangebot (Word Datei) in die Excel-Tabelle (Angebot_aktuell)'),
    'create-template': ('fileOfferService', 'create_template', 'Erstellt eine Excel-Vorlage (Angebot_aktuell)'),
    'create-delivery-agreement': ('deliveryAgreementService', 'create_delivery_agreement', 'Erstellt eine Word-Datei für die Ablieferungsvereinbarung'),
    'import-delivery-list-word': ('deliveryListService', 'import_delivery_list_word', 'Importiert die Ablieferungsvereinbarung und erstellt eine Excel-Datei für das Ablieferungsverzeichnis'),
    'import-delivery-list-excel': ('deliveryListService', 'import_delivery_list_excel', 'Importiert die Daten aus Angebot_aktuell und erstellt eine Excel-Datei für das Ablieferungsverzeichnis'),
}

def execute_command(command: str) -> None:
    """Führt das angegebene Kommando aus, indem die zugehörige Funktion importiert und aufgerufen wird."""
    if command in COMMANDS:
        file_name, function_name, _ = COMMANDS[command]
        try:
            module = importlib.import_module(file_name)  # Modul dynamisch importieren
            func = getattr(module, function_name)  # Funktion aus dem Modul holen
            func()  # Funktion ausführen
        except ModuleNotFoundError as e:
            messagebox.showerror("Fehler", f"Die Datei '{file_name}.py' wurde im services-Ordner nicht gefunden. Fehler: {e}")
        except AttributeError as e:
            messagebox.showerror("Fehler", f"Die Funktion '{function_name}' wurde in der Datei '{file_name}.py' nicht gefunden. Fehler: {e}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Es ist ein Fehler aufgetreten: {e}")
    else:
        messagebox.showerror("Fehler", f"Der Befehl '{command}' ist nicht im Mapping definiert.")

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
        execute_command(args.command)

    except Exception as e:
        messagebox.showerror("Fehler", f'Es ist ein Fehler aufgetreten: {e}')

# Startet das Skript, wenn es direkt ausgeführt wird
if __name__ == '__main__':
    main()