import os 
from tkinter import filedialog, messagebox, Tk
from typing import List, Tuple, Any
import config
import xlwings as xw
import win32com.client as win32comClient


def sheet_exists(sheet_name: str, workbook) -> bool:
    """Überprüft, ob ein Arbeitsblatt mit dem angegebenen Namen existiert."""
    try:
        workbook.sheets[sheet_name]
        return True
    except:
        return False


def import_file_offer_word() -> None:
    """Importiert ein Aktenangebotsformular und fügt die Daten in die Excel-Tabelle."""
    root = Tk()
    root.withdraw()

    try:
        app = xw.apps.active
        app.visible = False
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME]
    except Exception as e:
        messagebox.showerror("Excel Fehler", f"Es konnte keine Excel-Anwendung gestartet werden: {e}")
        return

    # Öffnen des Datei-Dialogs zur Auswahl einer Word-Datei
    file = filedialog.askopenfilename(filetypes=[("Word-Dokument", "*.docx")], title="Bitte Word-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return

    create_template()
    import_file_offer(file, workbook)
    messagebox.showinfo("Fertig", "Import von Aktenangebotsformular abgeschlossen.")


def import_file_offer(file_path: str, workbook) -> None:
    """Importiert die Daten aus einer Word-Datei in die Excel-Tabelle."""
    try:
        ws_target = workbook.sheets[config.CURRENT_OFFER_SHEET_NAME]
        copy_tables(file_path, ws_target)
    except Exception as e:
        messagebox.showerror("Fehler beim Import", f"Fehler beim Importieren der Datei: {e}")


def copy_tables(file_path: str, ws_target) -> None:
    """Kopiert die Tabellen aus einer Word-Datei in das Excel-Arbeitsblatt."""
    try:
        word_app = win32comClient.GetObject(None, "Word.Application") if not win32comClient.Dispatch("Word.Application") else win32comClient.Dispatch("Word.Application")
        word_app.Visible = False
        word_doc = word_app.Documents.Open(file_path, ReadOnly=True)
        direction, office, responsible, offer_date = read_meta_data_from_doc(word_doc)
        data_array = read_word_table(word_doc)
        write_data_to_sheet(ws_target, data_array, direction, office, responsible, offer_date)
        word_doc.Close(SaveChanges=False)
        word_app = None
        
    except Exception as e:
        messagebox.showerror("Word-Import Fehler", f"Fehler beim Lesen der Word-Datei: {e}")


def read_meta_data_from_doc(word_doc) -> Tuple[str, str, str, str]:
    """Liest die Metadaten aus einer Word-Datei."""
    try:
        full_text = word_doc.Content.text.replace('\r\n', '\r')
        meta_lines = full_text.split('\r')

        direction = safe_extract(meta_lines, 1, config.WORD_DIRECTION_TEXT)
        office = safe_extract(meta_lines, 2, config.WORD_OFFICE_TEXT)
        responsible = safe_extract(meta_lines, 3, config.WORD_RESPONSIBLE_TEXT)
        offer_date = safe_extract(meta_lines, 4, config.WORD_OFFER_DATE_TEXT)

        return direction, office, responsible, offer_date
    except Exception as e:
        messagebox.showerror("Fehler Word", f"Fehler beim Lesen von Metadaten: {e}")
        return "", "", "", ""


def safe_extract(arr: List[str], idx: int, tag: str) -> str:
    """Extrahiert sicher einen Text aus der Liste, basierend auf dem Index und Tag."""
    try:
        if 0 <= idx < len(arr):
            text = arr[idx].replace(tag, "")
            return clean_string(text)
    except:
        pass
    return ""


def read_word_table(word_doc) -> List[List[str]]:
    """Liest eine Tabelle aus einer Word-Datei und gibt die Daten als Liste zurück."""
    try:
        if word_doc.Tables.Count < config.WORD_TABLE_INDEX:
            return []
        
        word_table = word_doc.Tables(config.WORD_TABLE_INDEX)

        if word_table.Rows.Count < config.WORD_FIRST_DATA_ROW:
            return []
        
        tmp = []
        
        for r in range(config.WORD_FIRST_DATA_ROW, word_table.Rows.Count + 1):
            row = [safe_get_cell_text(word_table, r, c) for c in range(1, 6)]
            tmp.append(row)
        return tmp
    except Exception as e:
        messagebox.showerror("Fehler Word Table", f"Fehler beim Lesen der Word-Tabelle: {e}")
        return []


def safe_get_cell_text(word_table, r: int, c: int) -> str:
    """Sicheres Abrufen des Textes aus einer bestimmten Zelle der Word-Tabelle."""
    try:
        text = word_table.Cell(r, c).Range.text
        return clean_string(text)
    except:
        return ""


def clean_string(text: Any) -> str:
    """Bereinigt den Text von unnötigen Steuerzeichen und Leerzeichen."""
    if text is None or text == "":
        return ""
    
    s = str(text)
    s = s.replace('\r\n', ' ').replace('\r', ' ').replace('\n', ' ').replace('\t', ' ').replace('\x07', ' ').replace('\x0b', ' ').replace('\xa0', ' ')

    while '  ' in s:
        s = s.replace('  ', ' ')
    
    return s.strip()


def write_data_to_sheet(ws_target, data_array: List[List[str]], 
                        direction: str, office: str, responsible: str, offer_date: str) -> None:
    """Schreibt die extrahierten Daten in das Excel-Arbeitsblatt."""
    try:
        ws_target.range(f'O{config.ROW_BEHOERDE}').value = direction
        ws_target.range(f'O{config.ROW_AMT}').value = office
        ws_target.range(f'O{config.ROW_ZUSTAENDIG}').value = responsible
        ws_target.range(f'O{config.ROW_DATUM_ANGEBOT}').value = offer_date
        
        if not data_array:
            return
        
        out = []
        for row in data_array:
            if not any(row[1:5]):
                continue

            out.append([row[0], row[1], row[2] or "", row[3] or "", row[4].replace(",", ";") if row[4] else ""])
        
        if out:
            end_row = config.START_DATA_ROW + len(out) - 1
            ws_target.range(f'A{config.START_DATA_ROW}:E{end_row}').value = out
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Schreiben der Daten: {e}")


def create_template() -> None:
    """Erstellt eine Excel-Vorlage und fügt Dropdown-Listen hinzu."""
    try:
        app = xw.apps.active
        app.visible = False
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME]
        
        stammdaten_workbook = xw.Book(config.DATA_FILE_PATH)
        dateipfade_sheet = stammdaten_workbook.sheets[config.DATA_PATH_SHEET_NAME]
        select_fields = stammdaten_workbook.sheets[config.SELECTION_FIELDS_SHEET_NAME]
        
        select_values = select_fields.range(config.SELECTION_REASON_OFFER_RANGE).value
        conditional_values = [
            select_fields.range(config.SELECTION_COMPLETE_TAKEOVER_FIELDS).value,
            select_fields.range(config.SELECTION_PARTIAL_TAKEOVER_FIELDS).value,
            select_fields.range(config.SELECTION_NO_TAKEOVER_FIELDS).value,
            [select_fields.range(config.SELECTION_CANT_BE_OFFERED_FIELDS).value]
        ]

        # Holen des Vorlagenpfads
        template_path = dateipfade_sheet.range(config.DATA_PATH_FIELD_OFFER_TABLE).value
        stammdaten_workbook.close()

        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Der angegebene Pfad '{template_path}' ist ungültig.")
            return
        
        # Überprüfen, ob das Arbeitsblatt bereits existiert
        sheet_name = config.CURRENT_OFFER_SHEET_NAME
        if sheet_name in [sheet.name for sheet in workbook.sheets]:
            # Popup zur Bestätigung des Ersetzens anzeigen
            result = messagebox.askyesno("Bestätigung", f"Das Arbeitsblatt '{sheet_name}' existiert bereits. Möchten Sie es ersetzen?\n\nWarnung: Das Arbeitsblatt wird gelöscht und neu erstellt. Alle vorhandenen Daten gehen verloren.")
            if result:
                workbook.sheets[sheet_name].delete()
            else:
                return
        
        # Die Vorlage öffnen und das Arbeitsblatt kopieren
        template_workbook = xw.Book(template_path)
        template_sheet = template_workbook.sheets[sheet_name]
        template_sheet.copy(after=workbook.sheets[-1])
        template_workbook.close()

        # Dropdown-Listen hinzufügen
        add_dropdown_to_excel(workbook, select_values, conditional_values)
        
        # Speichern der Arbeitsmappe
        workbook.save()

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen des Templates: {e}")


def add_dropdown_to_excel(wb, values: list, conditional_values: list) -> None:
    """Fügt Dropdown-Listen in das Excel-Arbeitsblatt ein."""
    try:
        sheet = wb.sheets[config.CURRENT_OFFER_SHEET_NAME]
        g2_values_str = ";".join(values)
        cell_range_g2 = sheet.range(config.SELECTION_REASON_OFFER_LENGTH)

        cell_range_g2.api.Validation.Add(
            Type=3,
            AlertStyle=1,
            Operator=1,
            Formula1=g2_values_str
        )

        start_row = 11
        for i, cond_list in enumerate(conditional_values):
            if not cond_list:
                continue

            name = values[i].replace(" ", "_").replace("-", "_")
            vertical_list = [[val] for val in cond_list]
            end_row = start_row + len(vertical_list) - 1
            range_sheet = f"{config.SELECTION_SECTION}{start_row}:{config.SELECTION_SECTION}{end_row}"
            range_name = f"${config.SELECTION_SECTION}${start_row}:${config.SELECTION_SECTION}${end_row}"

            sheet.range(range_sheet).value = vertical_list
            try:
                if name in [n.name for n in wb.names]:
                    wb.names[name].delete()
                wb.names.add(name, f'={config.CURRENT_OFFER_SHEET_NAME}!{range_name}')
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Hinzufügen des benannten Bereichs '{name}': {e}")

            start_row = end_row + 1

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Hinzufügen der Dropdown-Liste: {e}")
