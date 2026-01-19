import os
import sys
from tkinter import filedialog, messagebox, Tk
from typing import List, Dict, Any

libs_path = os.path.join(os.path.dirname(__file__), '..', 'libs')

if os.path.exists(libs_path):
    sys.path.append(libs_path)
else:
    print(f"Warnung: Der Ordner '{libs_path}' wurde nicht gefunden. Bibliotheken aus diesem Ordner können nicht geladen werden.")

try:
    from openpyxl import Workbook, openpyxl
except ImportError as e:
    print(f"Fehler beim Importieren von openpyxl: {e}")

try:
    import docx
except ImportError:
    print("Warnung: docx nicht verfügbar.")

TARGET_SHEET_NAME = "Angebot_aktuell"
TEMPLATE_SHEET_NAME = "Vorlage"
WORD_TABLE_INDEX = 1
WORD_FIRST_DATA_ROW = 2

COL_INHALT = 1
COL_ZEITRAUM = 2
COL_LFM = 3
COL_GB = 4
COL_MEDIUM = 5

START_DATA_ROW = 6

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
    
    file = filedialog.askopenfilename(filetypes=[("Excel-Dateien", "*.xlsx")], title="Bitte Excel-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return
    doc_props, table_data = readExcelData(file)
    createDeliveryList(doc_props, table_data)


def createDeliveryList(doc_props: Dict[str, Any], table_data: List[List[str]]) -> None:
    try:
        wb = openpyxl.load_workbook("path_to_template.xlsx")
        ws = wb[TEMPLATE_SHEET_NAME]
        replace_placeholders_in_excel(ws, doc_props)
        insert_table_data(ws, table_data)
        
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Dateien", "*.xlsx")], title="Speicherort wählen...")
        if save_path:
            wb.save(save_path)
            messagebox.showinfo("Fertig", "Excel-Datei wurde erfolgreich erstellt.")
        else:
            messagebox.showinfo("Abbruch", "Speichern abgebrochen.")
    
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Excel-Datei: {e}")


def readDocxData(file_path: str) -> tuple[Dict[str, Any], List[List[str]]]:
    doc = docx.Document(file_path)
    
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
    for table in doc.tables:
        for row in table.rows:
            row_data = [cell.text.strip() for cell in row.cells]
            table_data.append(row_data)
    
    return doc_props, table_data


def readExcelData(file_path: str) -> tuple[Dict[str, Any], List[List[str]]]:
    wb = openpyxl.load_workbook(file_path)
    ws = wb["Angebot_aktuell"]
    
    doc_props = {
        "DirName": ws["O6"].value,
        "AmtName": ws["O7"].value,
        "AmtPName": ws["O8"].value,
        "ÜbernDatum": ws["O12"].value,
        "AngLfm": ws["O2"].value,
        "ÜbernLfm": ws["O3"].value,
        "AngGB": ws["O4"].value,
        "ÜbernGB": ws["O5"].value
    }
    
    table_data = []
    for row in ws.iter_rows(min_row=START_DATA_ROW, max_col=5, values_only=True):
        table_data.append(list(row))
    
    return doc_props, table_data


def replace_placeholders_in_excel(ws, doc_props: Dict[str, Any]) -> None:
    placeholders = ["<DirName>", "<AmtName>", "<AmtPName>", "<ÜbernDatum>", "<AngLfm>", "<AngGB>", "<ÜbernLfm>", "<ÜbernGB>"]
    values = [doc_props["DirName"], doc_props["AmtName"], doc_props["AmtPName"], doc_props["ÜbernDatum"],
              doc_props["AngLfm"], doc_props["AngGB"], doc_props["ÜbernLfm"], doc_props["ÜbernGB"]]
    
    for placeholder, value in zip(placeholders, values):
        for row in ws.iter_rows():
            for cell in row:
                if placeholder in str(cell.value):
                    cell.value = str(cell.value).replace(placeholder, value)


def insert_table_data(ws, table_data: List[List[str]]) -> None:
    row_start = 24
    for row in table_data:
        ws.append(row)
        row_start += 1
