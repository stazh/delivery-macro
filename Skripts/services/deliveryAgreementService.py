import os
from tkinter import messagebox
from datetime import datetime
import xlwings as xw
import shutil
import win32com.client as win32

def createDeliveryAgreement() -> None: 
    try: 
        app = xw.apps.active 
        app.visible = False 
        workbook = app.books['Ablieferungsmakro.xlsm'] 
        stammdaten_workbook = xw.Book('./Daten/Stammdaten.xlsx') 
        dateipfade_sheet = stammdaten_workbook.sheets['Dateipfade'] 
        template_path = dateipfade_sheet.range('B1').value 
        stammdaten_workbook.close() 
        
        if not template_path or not os.path.exists(template_path): 
            messagebox.showerror("Fehler", f"Der angegebene Pfad zur Vorlage '{template_path}' ist ungültig.") 
            return 
        
        data = importDeliveryAgreementExcel(workbook) 
        
        copy_path = os.path.abspath(f"Ablieferungsvereinbarung_{datetime.now().strftime('%d.%m.%Y')}.docx") 
        shutil.copy(template_path, copy_path) 
        
        if not os.path.exists(copy_path):
            messagebox.showerror("Fehler", f"Die Datei konnte nicht gefunden werden: {copy_path}")
            return

        word = win32.Dispatch("Word.Application") 
        word.visible = True
        doc = word.Documents.Open(copy_path)
        replacePlaceholdersInDoc(doc, data)
        output_path = os.path.abspath(f"Ablieferungsvereinbarung_{datetime.now().strftime('%d.%m.%Y')}.docx")
        doc.SaveAs(output_path)
        
        messagebox.showinfo("Fertig", f"Die Ablieferungsvereinbarung wurde erfolgreich erstellt und gespeichert.\nPfad: {output_path}")

    except Exception as e: 
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Ablieferungsvereinbarung: {e}") 
        print(f"Fehler beim Erstellen der Ablieferungsvereinbarung: {e}")


def importDeliveryAgreementExcel(workbook) -> dict:
    sheet = workbook.sheets['Angebot_aktuell']
    data = {
        "kuerz": str(sheet.range("O9").value).strip(),
        "amt_mail": str(sheet.range("Z6").value).strip(),
        "amt_zeichen": str(sheet.range("O7").value).strip(),
        "amt_pname": str(sheet.range("O8").value).strip(),
        "amt_name": str(sheet.range("O7").value).strip(),
        "bes_datum": str(sheet.range("O10").value).strip(),
        "ang_datum": str(sheet.range("O11").value).strip(),
        "uebern_datum": str(sheet.range("O12").value).strip(),
        "dir_name": str(sheet.range("O6").value).strip(),
        "ang_lfm": str(sheet.range("O2").value).strip(),
        "ang_gb": str(sheet.range("O4").value).strip(),
        "uebern_lfm": str(sheet.range("O3").value).strip(),
        "uebern_gb": str(sheet.range("O5").value).strip(),
    }
    return data


def replacePlaceholdersInDoc(doc, data: dict) -> None:
    try:
        replacements = {
            "<StAZHKürz>": data.get("kuerz", "<StAZHKürz>"),
            "<ErstellungsDatum>": datetime.now().strftime("%d.%m.%Y"),
            "<AmtZeichen>": data.get("amt_zeichen", "<AmtZeichen>"),
            "<AmtPName>": data.get("amt_pname", "<AmtPName>"),
            "<AmtMail>": data.get("amt_mail", "<AmtMail>"),
            "<BesDatum>": data.get("bes_datum", "<BesDatum>"),
            "<AblJahr>": datetime.now().year,
            "<AngDatum>": data.get("ang_datum", "<AngDatum>"),
            "<AngLfm>": data.get("ang_lfm", "<AngLfm>"),
            "<AngGB>": data.get("ang_gb", "<AngGB>"),
            "<ÜbernLfm>": data.get("uebern_lfm", "<ÜbernLfm>"),
            "<ÜbernGB>": data.get("uebern_gb", "<ÜbernGB>"),
            "<AmtName>": data.get("amt_name", "<AmtName>"),
            "<ÜbernDatum>": data.get("uebern_datum", "<ÜbernDatum>"),
            "<DirName>": data.get("dir_name", "<DirName>")
        }

        for placeholder, value in replacements.items():
            find_and_replace_in_word(doc, placeholder, value)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter in Word: {e}")
        print(f"Fehler beim Ersetzen der Platzhalter in Word: {e}")

#TODO: Look why it doesnt replace any placeholders in word document
def find_and_replace_in_word(doc, placeholder, value):
    try:
        range = doc.Content
        range.Find.ClearFormatting()
        range.Find.Text = placeholder
        range.Find.Replacement.Text = value
        range.Find.Execute(Replace=2)
        
        print(f"Ersetzt '{placeholder}' mit '{value}'")

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter: {e}")
        print(f"Fehler beim Ersetzen der Platzhalter: {e}")
