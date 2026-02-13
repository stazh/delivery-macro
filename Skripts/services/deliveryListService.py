import os
import sys
import re
from datetime import datetime
from tkinter import filedialog, messagebox, Tk
from typing import List, Dict, Any
import win32com.client as win32

try:
    import openpyxl
    import xlwings as xw
except ImportError as e:
    messagebox.showerror("Import-Fehler", f"Fehler beim Importieren der Bibliotheken: {e}")
    sys.exit(1)

TARGET_SHEET_NAME = "Angebot_aktuell"
WORD_TABLE_INDEX = 1
WORD_FIRST_DATA_ROW = 2

COL_INHALT = 1
COL_ZEITRAUM = 2
COL_LFM = 3
COL_GB = 4
COL_MEDIUM = 5

START_DATA_ROW = 2

def importDeliveryListWord() -> None:
    root = Tk()
    root.withdraw()
    
    file = filedialog.askopenfilename(filetypes=[("Word-Dokument", "*.docx")], title="Bitte Word-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return
    doc_props, table_data = readDocxData(file)
    createDeliveryList(doc_props, table_data)

def importDeliveryListExcel() -> None:
    root = Tk()
    root.withdraw()

    try:
        app = xw.apps.active
        workbook = app.books['Ablieferungsmakro.xlsm']
    except Exception as e:
        messagebox.showerror("Fehler", "Excel‑Instanz konnte nicht gefunden werden oder Datei ist nicht offen.")
        return


    try:
        doc_props, table_data = readExcelData(workbook)
        createDeliveryList(doc_props, table_data)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Verarbeiten der Abliefererdaten: {e}")

def createDeliveryList(doc_props: Dict[str, Any], table_data: List[List[str]]) -> None:
    try:
        stammdaten_workbook = xw.Book('./Daten/Stammdaten.xlsx')
        dateipfade_sheet = stammdaten_workbook.sheets['Dateipfade']
        
        template_rel = dateipfade_sheet.range('B2').value
        stammdaten_workbook.close()
        exc_file = os.path.abspath(template_rel)

        if not os.path.exists(exc_file):
            messagebox.showerror("Fehler", f"Die Vorlagen‑Datei wurde nicht gefunden:\n{exc_file}")
            return

        wb = openpyxl.load_workbook(exc_file)
        path = os.path.abspath("Ablieferungsverzeichnis.xlsx")
        wb.save(path)
        workbook = xw.Book('./Ablieferungsverzeichnis.xlsx')
        ws = workbook.sheets['Ablieferungsverzeichnis']
        replace_placeholders_in_excel(ws, doc_props)
        insert_table_data(ws, table_data)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Excel‑Datei: {e}")

def readDocxData(file_path: str) -> tuple[Dict[str, Any], List[List[str]]]:
    word = win32.DispatchEx("Word.Application") 
    word.visible = False
    word.Documents.Open(file_path)
    
    table_data = []
    raw_data = []
    doc = word.ActiveDocument
    table = doc.Tables[2]
    angLfm = 0
    angGB = 0

    for i in range(2, table.Rows.Count + 1):
        row = table.Rows(i)
        first_column_data = row.Cells(1).Range.Text.strip()
        third_column_data = row.Cells(3).Range.Text.strip()
        raw_data.append([first_column_data, third_column_data])

        lfm_match = re.search(r'(\d+(\.\d+)?)\s*Lfm', first_column_data)
        gb_match = re.search(r'(\d+(\.\d+)?)\s*GB', first_column_data)
        
        if lfm_match:
            angLfm += float(lfm_match.group(1))
        
        if gb_match:
            angGB += float(gb_match.group(1))

    doc_props = {
        "DirName": "<DirName>",
        "AmtName": "<AmtName>",
        "AmtPName": "<AmtPName>",
        "ÜbernDatum": "<ÜbernDatum>",
        "AngLfm": angLfm,
        "AngGB": angGB,
        "ÜbernLfm": "",
        "ÜbernGB": "",
        "AblNummer": "<AblNummer>"
    }

    table_data = convertWordTableData(raw_data)
    word.Quit()

    return doc_props, table_data

def convertWordTableData(raw_data: List[List[str]]) -> List[List[str]]:
    table_data = []
    for row in raw_data:
        inhalt = row[0]
        zeitraum = row[1]
        lfm = "0"
        gb = "0"
        medium = ""

        year_match = re.search(r',\s*(\d{4}(?:\s*[-+]\s*\d{4})*)', row[0])
        if year_match:
            year_index = year_match.start()
            parts = row[0][:year_index].strip(), row[0][year_index:].strip()
            inhalt = parts[0].removesuffix(',')
            zeitraum = parts[1].split(',')[1].strip()

            medium_match = re.search(r'\(([^)]+)\)(?!.*\()', row[0])
            if medium_match:
                medium = medium_match.group(1).strip()

        row[1] = row[1].strip().replace('\n', '').replace('\r', '').replace('\x07', '')
        if row[1].__contains__("/"):
            parts = row[1].split('/')
            lfm = parts[0]
            gb = parts[1]
        elif row[0].__contains__("Lfm"):
            lfm = row[1]
        else:
            gb = row[1]
        mark = "*" if gb != "0" else ""

        table_data.append([inhalt, zeitraum, lfm, gb, medium, mark])
    return table_data

def readExcelData(workbook) -> tuple[Dict[str, Any], List[List[str]]]:
    ws = workbook.sheets[TARGET_SHEET_NAME]

    übern_datum_raw = ws.range("O12").value or ""
    
    if übern_datum_raw is not None and übern_datum_raw != "":
        übern_datum = datetime.strftime(übern_datum_raw, "%d.%m.%Y")
    else:
        übern_datum = "<ÜbernDatum>"

    doc_props = {
        "DirName": ws.range("O6").value or "",
        "AmtName": ws.range("O7").value or "",
        "AmtPName": ws.range("O8").value or "",
        "ÜbernDatum": übern_datum,
        "AngLfm": "0" if ws.range("O2").value == 0.0 else str(ws.range("O2").value),
        "ÜbernLfm": "0" if ws.range("O3").value == 0.0 else str(ws.range("O3").value),
        "AngGB": "0" if ws.range("O4").value == 0.0 else str(ws.range("O4").value),
        "ÜbernGB": "0" if ws.range("O5").value == 0.0 else str(ws.range("O5").value),
        "AblNummer": ws.range("O13").value or ""
    }
    table_data = []
    rows = ws.range(f"A{START_DATA_ROW}:E{ws.cells.last_cell.row}").value
    accepted_data_row = ws.range(f"J{START_DATA_ROW}:K{ws.cells.last_cell.row}").value
    if rows:
        for row, accepted_row in zip(rows, accepted_data_row):
            if all(cell is None or str(cell).strip() == "" for cell in row):
                continue

            lfm_accepted = accepted_row[0] if accepted_row[0] is not None else "0"
            gb_accepted = accepted_row[1] if accepted_row[1] is not None else "0"
            row[2] = lfm_accepted
            row[3] = gb_accepted 

            if gb_accepted != "0":
                row.append("*")
            else:
                row.append("")

            table_data.append(row)
    return doc_props, table_data

def replace_placeholders_in_excel(ws, doc_props: Dict[str, Any]) -> None:
    placeholders = [
        "<DirName>", "<AmtName>", "<AmtPName>", "<ÜbernDatum>",
        "<AngLfm>", "<AngGB>", "<ÜbernLfm>", "<ÜbernGB>", "<AblNummer>"
    ]
    
    values = [
        str(doc_props["DirName"]),
        str(doc_props["AmtName"]),
        str(doc_props["AmtPName"]),
        str(doc_props["ÜbernDatum"]),
        str(doc_props["AngLfm"]),
        str(doc_props["AngGB"]),
        str(doc_props["ÜbernLfm"]),
        str(doc_props["ÜbernGB"]),
        str(doc_props["AblNummer"])
    ]

    for placeholder, value in zip(placeholders, values):
        for row in ws.range('A1:G20'):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if placeholder in cell.value:
                        cell.value = cell.value.replace(placeholder, value)


def insert_table_data(ws, table_data: List[List[str]]) -> None:
    row_start = 24
    for row in table_data:
        cell_range = ws.range(f'B{row_start}:G{row_start}')
        cell_range.value = row
        row_start += 1
