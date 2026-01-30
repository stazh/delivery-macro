import os
import sys
from tkinter import filedialog, messagebox, Tk
from typing import List, Dict, Any

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
        temp_path = os.path.abspath("Ablieferungsverzeichnis.xlsx")
        wb.save(temp_path)
        workbook = xw.Book('./Ablieferungsverzeichnis.xlsx')
        ws = workbook.sheets['Ablieferungsverzeichnis']
        replace_placeholders_in_excel(ws, doc_props)
        insert_table_data(ws, table_data)

        app = xw.App(visible=True)
        app.books.open(temp_path)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Excel‑Datei: {e}")




def readDocxData(file_path: str) -> tuple[Dict[str, Any], List[List[str]]]:
    #doc = docx.Document(file_path)
    
    doc_props = {
        "DirName": "Beispiel Direktion",
        "AmtName": "Beispiel Amt",
        "AmtPName": "Beispiel Amtsperson",
        "ÜbernDatum": "01.01.2023",
        "AngLfm": "100",
        "ÜbernLfm": "200",
        "AngGB": "500",
        "ÜbernGB": "1000"
    }
    
    table_data = []
    #for table in doc.tables:
        #for row in table.rows:
            #row_data = [cell.text.strip() for cell in row.cells]
            #table_data.append(row_data)
    
    return doc_props, table_data


def readExcelData(workbook) -> tuple[Dict[str, Any], List[List[str]]]:
    ws = workbook.sheets[TARGET_SHEET_NAME]

    doc_props = {
        "DirName": ws.range("O6").value or "",
        "AmtName": ws.range("O7").value or "",
        "AmtPName": ws.range("O8").value or "",
        "ÜbernDatum": ws.range("O12").value or "",
        "AngLfm": ws.range("O2").value or "",
        "ÜbernLfm": ws.range("O3").value or "",
        "AngGB": ws.range("O4").value or "",
        "ÜbernGB": ws.range("O5").value or ""
    }

    table_data = []
    rows = ws.range(f"A{START_DATA_ROW}:E{ws.cells.last_cell.row}").value

    if rows:
        for row in rows:
            if all(cell is None or str(cell).strip() == "" for cell in row):
                continue
            table_data.append(row)

    return doc_props, table_data



def replace_placeholders_in_excel(ws, doc_props: Dict[str, Any]) -> None:
    placeholders = [
        "<DirName>", "<AmtName>", "<AmtPName>", "<ÜbernDatum>",
        "<AngLfm>", "<AngGB>", "<ÜbernLfm>", "<ÜbernGB>"
    ]
    
    values = [
        str(doc_props["DirName"]),
        str(doc_props["AmtName"]),
        str(doc_props["AmtPName"]),
        str(doc_props["ÜbernDatum"]),
        str(doc_props["AngLfm"]),
        str(doc_props["AngGB"]),
        str(doc_props["ÜbernLfm"]),
        str(doc_props["ÜbernGB"])
    ]

    for placeholder, value in zip(placeholders, values):
        for row in ws.iter_rows():
            for cell in row:
                cell_val = str(cell.value) if cell.value is not None else ""
                if placeholder in cell_val:
                    cell.value = cell_val.replace(placeholder, value)



def insert_table_data(ws, table_data: List[List[str]]) -> None:
    row_start = 24
    for row in table_data:
        ws.append(row)
        row_start += 1
