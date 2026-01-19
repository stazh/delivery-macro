import sys
import os
import re
from tkinter import filedialog, messagebox, Tk
from typing import Optional, List, Tuple, Dict, Any

libs_path = os.path.join(os.getcwd(), 'libs')

if os.path.exists(libs_path):
    sys.path.append(libs_path)
else:
    print(f"Warnung: Der Ordner '{libs_path}' wurde nicht gefunden. Bibliotheken aus diesem Ordner können nicht geladen werden.")

try:
    import xlwings as xw
except ImportError as e:
    print(f"Fehler beim Importieren von xlwings: {e}")

try:
    from win32com.client import GetObject, Dispatch
except ImportError:
    print("Warnung: win32com nicht verfügbar. Word-Integration kann nicht verwendet werden.")

TARGET_SHEET_NAME = "Angebot_aktuell"
TEMPLATE_SHEET_NAME = "Vorlage"
STEUERBOARD_SHEET_NAME = "Steuerboard"
WORD_TABLE_INDEX = 1
WORD_FIRST_DATA_ROW = 2

COL_INHALT = 1
COL_ZEITRAUM = 2
COL_LFM = 3
COL_GB = 4
COL_MEDIUM = 5

ROW_BEHOERDE = 1
ROW_AMT = 2
ROW_ZUSTAENDIG = 3
ROW_DATUM_ANGEBOT = 4
COL_METADATA = 2

START_DATA_ROW = 6

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
        app = xw.App(visible=True)
        workbook = app.books.active
    except Exception as e:
        print(f"Fehler beim Starten von Excel: {e}")
        messagebox.showerror("Excel Fehler", "Es konnte keine Excel-Anwendung gestartet werden.")
        return

    file = filedialog.askopenfilename(filetypes=[("Word-Dokument", "*.docx")], title="Bitte Word-Datei auswählen...")
    if not file:
        messagebox.showinfo("Abbruch", "Kein Dateipfad ausgewählt. Der Vorgang wird abgebrochen.")
        return

    if sheetExists(TARGET_SHEET_NAME, workbook):
        deleteTable = messagebox.askquestion(
            "Tabellenblatt existiert",
            f"Das Tabellenblatt '{TARGET_SHEET_NAME}' existiert bereits.\nSoll es gelöscht werden?",
            icon='question', 
            type='yesnocancel'
        )
        
        if deleteTable == 'yes':
            saveTable = messagebox.askquestion(
                "Vorher speichern?", 
                f"Möchten Sie das aktuelle '{TARGET_SHEET_NAME}'-Blatt speichern, bevor es gelöscht wird?",
                icon='question', 
                type='yesnocancel'
            )

            if saveTable == 'yes':
                saveFilePath = filedialog.asksaveasfilename(defaultextension=".xlsm", filetypes=[("Excel Macro-Enabled Workbook", "*.xlsm")], title="Speicherort wählen...")
                if not saveFilePath:
                    messagebox.showinfo("Abbruch", "Speichern abgebrochen. Import wird nicht durchgeführt.")
                    return

                workbook.sheets[TARGET_SHEET_NAME].copy()
                newWorkbook = app.books.active
                newWorkbook.save(saveFilePath)
                newWorkbook.close()

                workbook.sheets[TARGET_SHEET_NAME].delete()

            elif saveTable == 'no':
                workbook.sheets[TARGET_SHEET_NAME].delete()

            elif saveTable == 'cancel':
                messagebox.showinfo("Abbruch", "Löschen abgebrochen. Import wird nicht durchgeführt.")
                return
        elif deleteTable == 'no':
            pass
        elif deleteTable == 'cancel':
            messagebox.showinfo("Abbruch", "Import abgebrochen. Das Tabellenblatt bleibt unverändert.")
            return

    importFileOffer(file, workbook)

    messagebox.showinfo("Information", "Import abgeschlossen. Bitte die rot markierten Felder ausfüllen und die hellblauen Felder prüfen.")
    app.quit()


def importFileOffer(file_path: str, workbook) -> None:
    try:
        template_sheet = workbook.sheets[TEMPLATE_SHEET_NAME]
        steuerboard_index = workbook.sheets[STEUERBOARD_SHEET_NAME].index
        
        template_sheet.copy(before=workbook.sheets[STEUERBOARD_SHEET_NAME])
        
        ws_target = workbook.sheets[steuerboard_index - 1]
        ws_target.name = TARGET_SHEET_NAME
        
        copyTables(file_path, ws_target)
        formatTitleRows(ws_target)
        
    except Exception as e:
        messagebox.showerror("Fehler beim Import", f"Fehler beim Importieren der Datei: {e}")


def copyTables(file_path: str, ws_target) -> None:
    try:
        try:
            word_app = GetObject(None, "Word.Application")
        except:
            word_app = Dispatch("Word.Application")
        
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
        print(f"Fehler beim Lesen von Metadaten: {e}")
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
            for c in range(1, 6):  # 5 Spalten
                cell_text = safeGetCellText(word_table, r, c)
                row.append(cell_text)
            tmp.append(row)
        
        return tmp
    except Exception as e:
        print(f"Fehler beim Lesen der Word-Tabelle: {e}")
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
        ws_target.range(f'B{ROW_BEHOERDE}').value = direction
        ws_target.range(f'B{ROW_AMT}').value = office
        ws_target.range(f'B{ROW_ZUSTAENDIG}').value = responsible
        ws_target.range(f'B{ROW_DATUM_ANGEBOT}').value = offer_date
        
        if not data_array:
            return
        
        out = []
        for row in data_array:
            if not any(row[1:5]):
                out.append([row[0], "", "", "", ""]) 
            else:
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
        print(f"Fehler bei Einheiten-Konvertierung: {e}")
        return ""


def formatTitleRows(ws: Any, first_data_row: int = START_DATA_ROW) -> None:
    try:
        last_row = ws.cells.last_cell.row
        
        for r in range(first_data_row, last_row + 1):
            check_range = ws.range(f'B{r}:E{r}')
            
            non_empty_count = 0
            for cell in check_range:
                if cell.value:
                    non_empty_count += 1
            
            if non_empty_count == 0:
                title_color = ws.range(f'A{r}').color
                
                for c in range(1, 13):
                    if c not in [6, 8]:
                        ws.cells(r, c).color = title_color
    
    except Exception as e:
        print(f"Fehler beim Formatieren: {e}")

def createTemplate() -> None:
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Create Template aufgerufen", "Dies ist eine einfache Nachricht.")

    try:
        app = xw.App(visible=True)
        workbook = app.books.active
    except Exception as e:
        print(f"Fehler beim Starten von Excel: {e}")
        messagebox.showerror("Excel Fehler", "Es konnte keine Excel-Anwendung gestartet werden.")
        return

    if sheetExists(TARGET_SHEET_NAME, workbook):
        deleteTable = messagebox.askquestion(
            f"Blatt '{TARGET_SHEET_NAME}' gefunden",
            f"Das Tabellenblatt '{TARGET_SHEET_NAME}' existiert bereits.\n"
            f"Soll es gelöscht werden, um ein neues leeres Blatt zu erstellen?",
            icon='question', 
            type='yesnocancel'
        )
        
        if deleteTable == 'yes':
            saveTable = messagebox.askquestion(
                "Vorher speichern?", 
                f"Möchten Sie das aktuelle '{TARGET_SHEET_NAME}'-Blatt speichern, bevor es gelöscht wird?",
                icon='question', 
                type='yesnocancel'
            )

            if saveTable == 'yes':
                saveFilePath = filedialog.asksaveasfilename(
                    initialfile=f"{TARGET_SHEET_NAME}.xlsm",
                    defaultextension=".xlsm", 
                    filetypes=[("Excel Macro-Enabled Workbook (*.xlsm)", "*.xlsm")], 
                    title="Bitte Speicherort und Dateiname wählen..."
                )
                
                if not saveFilePath:
                    messagebox.showinfo("Abbruch", "Speichervorgang abgebrochen. Vorgang wird nicht ausgeführt.")
                    app.quit()
                    return

                workbook.sheets[TARGET_SHEET_NAME].copy()
                newWorkbook = app.books.active
                newWorkbook.save(saveFilePath)
                newWorkbook.close()

                workbook.sheets[TARGET_SHEET_NAME].delete()

            elif saveTable == 'no':
                workbook.sheets[TARGET_SHEET_NAME].delete()

            elif saveTable == 'cancel':
                messagebox.showinfo("Abbruch", "Löschen abgebrochen. Kein neues Blatt wird erstellt.")
                app.quit()
                return
                
        elif deleteTable == 'no':
            messagebox.showinfo("Information", "Das Blatt bleibt erhalten. Es wird kein neues leeres Blatt erstellt.")
            app.quit()
            return
            
        elif deleteTable == 'cancel':
            messagebox.showinfo("Abbruch", "Vorgang abgebrochen.")
            app.quit()
            return

    createEmptyTemplate(workbook)
    
    messagebox.showinfo("Fertig", f"Neues leeres '{TARGET_SHEET_NAME}'-Blatt wurde erstellt.")
    app.quit()


def createEmptyTemplate(workbook) -> None:
    try:
        workbook.sheets.add(name=TARGET_SHEET_NAME)
    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen des neuen Blattes: {e}")