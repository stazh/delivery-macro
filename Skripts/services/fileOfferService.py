import sys
import os
from tkinter import filedialog, messagebox, Tk
from typing import List, Tuple, Any
import config
import xlwings as xw
import win32com.client as win32comClient


def sheetExists(sheetName, workbook) -> bool:
    try:
        workbook.sheets[sheetName]
        return True
    except:
        return False

def importFileOfferWord() -> None:
    root = Tk()
    root.withdraw()

    try:
        app = xw.apps.active
        app.visible = False
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME] 
    except Exception as e:
        messagebox.showerror("Excel Fehler", f"Es konnte keine Excel-Anwendung gestartet werden: {e}")
        return

    file = filedialog.askopenfilename(filetypes=[("Word-Dokument", "*.docx")], title="Bitte Word-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return

    createTemplate()
    importFileOffer(file, workbook)
    messagebox.showinfo("Fertig", "Import von Aktenangebotsformular abgeschlossen.")

def importFileOffer(file_path: str, workbook) -> None:
    try:
        ws_target = workbook.sheets[config.CURRENT_OFFER_SHEET_NAME]
        copyTables(file_path, ws_target)
    except Exception as e:
        messagebox.showerror("Fehler beim Import", f"Fehler beim Importieren der Datei: {e}")

def copyTables(file_path: str, ws_target) -> None:
    try:
        try:
            word_app = win32comClient.GetObject(None, "Word.Application")
        except:
            word_app = win32comClient.Dispatch("Word.Application")
        
        word_doc = word_app.Documents.Open(file_path, ReadOnly=True)
        direction, office, responsible, offer_date = readMetaDataFromDoc(word_doc)
        data_array = readWordTable(word_doc)
        writeDataToSheet(ws_target, data_array, direction, office, responsible, offer_date)
        word_doc.Close(SaveChanges=False)
        word_app = None
        
    except Exception as e:
        messagebox.showerror("Word-Import Fehler", f"Fehler beim Lesen der Word-Datei: {e}")

def readMetaDataFromDoc(word_doc) -> Tuple[str, str, str, str]:
    try:
        full_text = word_doc.Content.text.replace('\r\n', '\r')
        meta_lines = full_text.split('\r')

        direction = safeExtract(meta_lines, 1, config.WORD_DIRECTION_TEXT)
        office = safeExtract(meta_lines, 2, config.WORD_OFFICE_TEXT)
        responsible = safeExtract(meta_lines, 3, config.WORD_RESPONSIBLE_TEXT)
        offer_date = safeExtract(meta_lines, 4, config.WORD_OFFER_DATE_TEXT)

        return direction, office, responsible, offer_date
    except Exception as e:
        messagebox.showerror("Fehler Word", f"Fehler beim Lesen von Metadaten: {e}")
        return "", "", "", ""

def safeExtract(arr: List[str], idx: int, tag: str) -> str:
    try:
        if 0 <= idx < len(arr):
            text = arr[idx].replace(tag, "")
            return cleanString(text)
    except:
        pass
    return ""

def readWordTable(word_doc) -> List[List[str]]:
    try:
        if word_doc.Tables.Count < config.WORD_TABLE_INDEX:
            return []
        
        word_table = word_doc.Tables(config.WORD_TABLE_INDEX)

        if word_table.Rows.Count < config.WORD_FIRST_DATA_ROW:
            return []
        
        data_rows = word_table.Rows.Count - (config.WORD_FIRST_DATA_ROW - 1)
        tmp = []
        
        for r in range(config.WORD_FIRST_DATA_ROW, word_table.Rows.Count + 1):
            row = []
            for c in range(1, 6):
                cell_text = safeGetCellText(word_table, r, c)
                row.append(cell_text)
            tmp.append(row)
        return tmp
    except Exception as e:
        messagebox.showerror("Fehler Word Table", f"Fehler beim Lesen der Word-Tabelle: {e}")
        return []

def safeGetCellText(word_table, r: int, c: int) -> str:
    try:
        text = word_table.Cell(r, c).Range.text
        return cleanString(text)
    except:
        return ""

def cleanString(text: Any) -> str:
    if text is None or text == "":
        return ""
    
    s = str(text)
    s = s.replace('\r\n', ' ')
    s = s.replace('\r', ' ')
    s = s.replace('\n', ' ')
    s = s.replace('\t', ' ')
    s = s.replace('\x07', ' ')
    s = s.replace('\x0b', ' ')
    s = s.replace('\xa0', ' ')
    
    while '  ' in s:
        s = s.replace('  ', ' ')
    
    return s.strip()

def writeDataToSheet(ws_target, data_array: List[List[str]], 
                    direction: str, office: str, responsible: str, offer_date: str) -> None:
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

            lfm_value = ""
            if row[2] and row[2].strip():
                lfm_value = row[2]
            
            gb_value = ""
            if row[3] and row[3].strip():
                gb_value = row[3]
            
            medium_value = row[4].replace(",", ";") if row[4] else ""
            
            out.append([row[0], row[1], lfm_value, gb_value, medium_value])
        
        if out:
            end_row = config.START_DATA_ROW + len(out) - 1
            ws_target.range(f'A{config.START_DATA_ROW}:E{end_row}').value = out
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Schreiben der Daten: {e}")

def createTemplate() -> None:
    try:
        app = xw.apps.active
        app.visible = False
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME]
        
        stammdaten_workbook = xw.Book(config.DATA_FILE_PATH)
        dateipfade_sheet = stammdaten_workbook.sheets[config.DATA_PATH_SHEET_NAME]
        select_fields = stammdaten_workbook.sheets[config.SELECTION_FIELDS_SHEET_NAME]
        
        select_values = select_fields.range(config.SELECTION_REASON_OFFER_RANGE).value

        complete_takeover_list = select_fields.range(config.SELECTION_COMPLETE_TAKEOVER_FIELDS).value
        partial_takeover_list = select_fields.range(config.SELECTION_PARTIAL_TAKEOVER_FIELDS).value
        no_takeover_list = select_fields.range(config.SELECTION_NO_TAKEOVER_FIELDS).value
        cant_be_offered_list = select_fields.range(config.SELECTION_CANT_BE_OFFERED_FIELDS).value
        conditional_values = [complete_takeover_list, partial_takeover_list, no_takeover_list, [cant_be_offered_list]]

        template_path = dateipfade_sheet.range(config.DATA_PATH_FIELD_OFFER_TABLE).value
        stammdaten_workbook.close()
        
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Der angegebene Pfad '{template_path}' ist ungültig.")
            return
        
        template_workbook = xw.Book(template_path)
        template_sheet = template_workbook.sheets[config.CURRENT_OFFER_SHEET_NAME]

        template_sheet.copy(after=workbook.sheets[-1])
        
        template_workbook.close()

        add_dropdown_to_excel(workbook, select_values, conditional_values)
        workbook.save()

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen des Templates: {e}")

def add_dropdown_to_excel(wb, values: list, conditional_values: list) -> None:
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
