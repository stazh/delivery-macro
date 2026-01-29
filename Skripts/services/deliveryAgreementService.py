import sys
import os
from tkinter import filedialog, messagebox, Tk
from typing import List, Dict, Any
from datetime import datetime

try:
    import openpyxl
    import docx
except ImportError as e:
    messagebox.showerror("Import-Fehler", f"Fehler beim Importieren der Bibliotheken: {e}")
    sys.exit(1)

def importDeliveryAgreementExcel() -> Dict[str, Any]:
    base_dir = os.path.dirname(__file__)
    excel_file = os.path.join(base_dir, "Ablieferungsmakro.xlsm")

    wb = openpyxl.load_workbook(excel_file, keep_vba=True)
    sheet = wb["Angebot_aktuell"]

    data = {
        "kuerz": str(sheet["O9"].value).strip(),
        "amt_mail": str(sheet["Z6"].value).strip(),
        "amt_zeichen": str(sheet["O7"].value).strip(),
        "amt_pname": str(sheet["O8"].value).strip(),
        "amt_name": str(sheet["O7"].value).strip(),
        "bes_datum": str(sheet["O10"].value).strip(),
        "ang_datum": str(sheet["O11"].value).strip(),
        "uebern_datum": str(sheet["O12"].value).strip(),
        "dir_name": str(sheet["O6"].value).strip(),
        "ang_lfm": str(sheet["O2"].value).strip(),
        "ang_gb": str(sheet["O4"].value).strip(),
        "uebern_lfm": str(sheet["O3"].value).strip(),
        "uebern_gb": str(sheet["O5"].value).strip(),
    }

    return data


def createDeliveryAgreement(data: Dict[str, Any]) -> None:
    base_dir = os.path.dirname(__file__)
    docx_file = os.path.join(base_dir, "/Vorlagen/Ablieferungsvereinbarung, Vorlage.dotx")
    doc = docx.Document(docx_file)

    replacePlaceholdersInWord(doc, data)
    fillTablesInWord(doc, data)
    doc.save("Ablieferungsvereinbarung.docx")


def replacePlaceholdersInWord(doc: docx.Document, data: Dict[str, Any]) -> None:
    placeholders = [
        "<StAZHKürz>", "<StAZHName>", "<StAZHNummer>", "<StAZHMail>", "<ErstellungsDatum>",
        "<AmtZeichen>", "<AmtPName>", "<AmtMail>", "<BesDatum>", "<AblJahr>",
        "<AngDatum>", "<AngLfm>", "<AngGB>", "<ÜbernLfm>", "<ÜbernGB>",
        "<AmtName>", "<ÜbernDatum>", "<DirName>"
    ]
    values = [
        data["kuerz"], data["amt_name"], data["amt_zeichen"], data["amt_mail"], datetime.now().strftime("%d.%m.%Y"),
        data["amt_zeichen"], data["amt_pname"], data["amt_mail"], data["bes_datum"], datetime.now().year,
        data["ang_datum"], data["ang_lfm"], data["ang_gb"], data["uebern_lfm"], data["uebern_gb"],
        data["amt_name"], data["uebern_datum"], data["dir_name"]
    ]
    
    for p, v in zip(placeholders, values):
        for para in doc.paragraphs:
            if p in para.text:
                inline = para.runs
                for run in inline:
                    if p in run.text:
                        run.text = run.text.replace(p, v)
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    if p in cell.text:
                        cell.text = cell.text.replace(p, v)


def fillTablesInWord(doc: docx.Document, data: Dict[str, Any]) -> None:
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if "Aktengruppe" in cell.text:
                    cell.text = f"Beispielwert für {data['ang_lfm']}"
                elif "Menge" in cell.text:
                    cell.text = f"Beispielmenge für {data['uebern_lfm']}"

data = importDeliveryAgreementExcel()
createDeliveryAgreement(data)
