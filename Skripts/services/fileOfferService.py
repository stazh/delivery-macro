import sys
import os
import re
from tkinter import filedialog, messagebox, Tk
from typing import List, Tuple, Any

try:
    import xlwings as xw
    import win32com.client as win32comClient
except ImportError as e:
    messagebox.showerror("Import-Fehler", f"Fehler beim Importieren der Bibliotheken: {e}")
    sys.exit(1)

TARGET_SHEET_NAME = "Angebot_aktuell"
WORD_TABLE_INDEX = 2
WORD_FIRST_DATA_ROW = 3

COL_INHALT = 1
COL_ZEITRAUM = 2
COL_LFM = 3
COL_GB = 4
COL_MEDIUM = 5

ROW_BEHOERDE = 6
ROW_AMT = 7
ROW_ZUSTAENDIG = 8
ROW_DATUM_ANGEBOT = 11
COL_METADATA = 2

START_DATA_ROW = 2

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
        workbook = app.books['Ablieferungsmakro.xlsm']
    except Exception as e:
        messagebox.showerror("Excel Fehler", "Es konnte keine Excel-Anwendung gestartet werden.")
        return

    file = filedialog.askopenfilename(filetypes=[("Word-Dokument", "*.docx")], title="Bitte Word-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return

    createTemplate()
    importFileOffer(file, workbook)
    messagebox.showinfo("Information", "Import von Aktenangebotsformular abgeschlossen.")


def importFileOffer(file_path: str, workbook) -> None:
    try:
        ws_target = workbook.sheets[TARGET_SHEET_NAME]
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

        direction = safeExtract(meta_lines, 1, "Direktion (Behörde):")
        office = safeExtract(meta_lines, 2, "Amtsstelle:")
        responsible = safeExtract(meta_lines, 3, "Zuständig in der Amtsstelle:")
        offer_date = safeExtract(meta_lines, 4, "Datum des Angebots:")

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
        if word_doc.Tables.Count < WORD_TABLE_INDEX:
            return []
        
        word_table = word_doc.Tables(WORD_TABLE_INDEX)

        if word_table.Rows.Count < WORD_FIRST_DATA_ROW:
            return []
        
        data_rows = word_table.Rows.Count - (WORD_FIRST_DATA_ROW - 1)
        tmp = []
        
        for r in range(WORD_FIRST_DATA_ROW, word_table.Rows.Count + 1):
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
        ws_target.range(f'O{ROW_BEHOERDE}').value = direction
        ws_target.range(f'O{ROW_AMT}').value = office
        ws_target.range(f'O{ROW_ZUSTAENDIG}').value = responsible
        ws_target.range(f'O{ROW_DATUM_ANGEBOT}').value = offer_date
        if not data_array:
            return
        
        out = []
        for row in data_array:
            if not any(row[1:5]):
                continue

            lfm_value = ""
            if row[2] and row[2].strip():
                lfm_value = convertToLfm(row[2])
            
            gb_value = ""
            if row[3] and row[3].strip():
                gb_value = convertToGb(row[3])
            
            medium_value = row[4].replace(",", ";") if row[4] else ""
            
            out.append([
                row[0],
                row[1],
                lfm_value,
                gb_value,
                medium_value
            ])
        
        if out:
            end_row = START_DATA_ROW + len(out) - 1
            ws_target.range(f'A{START_DATA_ROW}:E{end_row}').value = out
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Schreiben der Daten: {e}")



def convertToLfm(value: Any) -> float:
    return convertWithUnits(value, "LENGTH")


def convertToGb(value: Any) -> float:
    return convertWithUnits(value, "SIZE")


def convertWithUnits(value: Any, unit_type: str) -> float:
    try:
        if value is None or value == "":
            return ""
        
        s = str(value).lower().replace(",", ".").replace(" ", "")
        if s == "":
            return ""
        
        match = re.match(r'([0-9.]+)(.*)', s)
        if not match:
            return ""
        
        num_str = match.group(1)
        unit = match.group(2)
        
        try:
            num = float(num_str)
        except:
            return ""
        
        factor = 1.0
        
        if unit_type == "LENGTH":
            if unit in ["", "m", "lfm", "lm", "meter"]:
                factor = 1.0
            elif unit == "cm":
                factor = 0.01
            elif unit == "mm":
                factor = 0.001
            elif unit == "km":
                factor = 1000.0
        
        elif unit_type == "SIZE":
            if unit in ["", "g", "gb", "gigabyte"]:
                factor = 1.0
            elif unit == "tb":
                factor = 1024.0
            elif unit == "mb":
                factor = 1.0 / 1024.0
            elif unit in ["kb", "k"]:
                factor = 1.0 / 1048576.0
        
        return num * factor
    
    except Exception as e:
        messagebox.showerror("Fehler Convert", f"Fehler bei Einheiten-Konvertierung: {e}")
        return ""


def createTemplate() -> None:
    try:
        app = xw.apps.active
        app.visible = False
        workbook = app.books['Ablieferungsmakro.xlsm']
        
        stammdaten_workbook = xw.Book('./Daten/Stammdaten.xlsx')
        dateipfade_sheet = stammdaten_workbook.sheets['Dateipfade']
        select_fields = stammdaten_workbook.sheets['Auswahlfelder']
        
        select_values = select_fields.range('A1:D1').value

        complete_takeover_list = select_fields.range('A2:A11').value
        partial_takeover_list = select_fields.range('B2:B5').value
        no_takeover_list = select_fields.range('C2:C9').value
        cant_be_offered_list = select_fields.range('D2').value
        conditional_values = [complete_takeover_list, partial_takeover_list, no_takeover_list, cant_be_offered_list]

        template_path = dateipfade_sheet.range('B4').value
        stammdaten_workbook.close()
        
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Der angegebene Pfad '{template_path}' ist ungültig.")
            return
        
        template_workbook = xw.Book(template_path)
        template_sheet = template_workbook.sheets[TARGET_SHEET_NAME]

        template_sheet.copy(after=workbook.sheets[-1])
        
        template_workbook.close()

        add_dropdown_to_excel(workbook, select_values, conditional_values)
        workbook.save()

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen des Templates: {e}")


def add_dropdown_to_excel(wb, values: list, conditional_values: list) -> None:
    try:
        sheet = wb.sheets["Angebot_aktuell"]
        g2_values_str = ";".join(values)
        cell_range_g2 = sheet.range('G2:G6268')
        cell_range_h2 = sheet.range('H2:H6268')

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
            range_sheet = f"AZ{start_row}:AZ{end_row}"
            sheet.range(range_sheet).value = vertical_list
            range_address = sheet.range(range_sheet)

            try:
                if name in [n.name for n in wb.names]:
                    wb.names[name].delete()
                wb.names.add(name, range_address) #TODO: Prüfe, warum der Bereich nicht hinzugefügt wird und es feststeckt (keine Fehlermeldung erhalten)
            except Exception as e:
                messagebox.showerror("Fehler", f"Fehler beim Hinzufügen des benannten Bereichs '{name}': {e}")

            start_row = end_row + 1

        formula = "=INDIREKT(WECHSELN(G2;\" \";\"_\"))"

        cell_range_h2.api.Validation.Add(
            Type=3,
            AlertStyle=1,
            Operator=1,
            Formula1=formula
        )

        messagebox.showinfo("Erfolg", "Dropdowns wurden erfolgreich hinzugefügt.")

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Hinzufügen der Dropdown-Liste: {e}")
