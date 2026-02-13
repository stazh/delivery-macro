import os
import re
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
        contact_sheet = stammdaten_workbook.sheets['Kontakte'] 
        template_path = dateipfade_sheet.range('B1').value 
        
        if not template_path or not os.path.exists(template_path): 
            messagebox.showerror("Fehler", f"Der angegebene Pfad zur Vorlage '{template_path}' ist ungültig.") 
            return 
        
        data, table_data, declined_table_data = importDeliveryAgreementExcel(workbook)
        
        wordPath = os.path.abspath(f"Ablieferungsvereinbarung_{datetime.now().strftime('%d.%m.%Y')}.docx") 
        shutil.copy(template_path, wordPath) 
        
        if not os.path.exists(wordPath):
            messagebox.showerror("Fehler", f"Die Datei konnte nicht gefunden werden: {wordPath}")
            return

        word = win32.DispatchEx("Word.Application") 
        word.visible = True
        word.Documents.Open(wordPath)
        replacePlaceholdersInDoc(word, data, contact_sheet)
        insertTableDataIntoWord(word, table_data, declined_table_data)
        stammdaten_workbook.close() 
        word.ActiveDocument.Save()
        
        messagebox.showinfo("Fertig", f"Die Ablieferungsvereinbarung wurde erfolgreich erstellt und gespeichert.\nPfad: {wordPath}")

    except Exception as e: 
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Ablieferungsvereinbarung: {e}") 


def importDeliveryAgreementExcel(workbook) -> dict:
    sheet = workbook.sheets['Angebot_aktuell']
    
    data = {
        "kuerz": sheet.range("O9").value if sheet.range("O9").value else "<StAZHKürz>",
        "amt_mail": sheet.range("Z6").value if sheet.range("Z6").value else "<AmtMail>",
        "amt_zeichen": sheet.range("O7").value if sheet.range("O7").value else "<AmtZeichen>",
        "amt_pname": sheet.range("O8").value if sheet.range("O8").value else "<AmtPName>",
        "amt_name": sheet.range("O7").value if sheet.range("O7").value else "<AmtName>",
        "bes_datum": sheet.range("O10").value if sheet.range("O10").value else "<BesDatum>",
        "ang_datum": sheet.range("O11").value if sheet.range("O11").value else "<AngDatum>",
        "uebern_datum": sheet.range("O12").value if sheet.range("O12").value else "<ÜbernDatum>",
        "dir_name": sheet.range("O6").value if sheet.range("O6").value else "<DirName>",
        "ang_lfm": sheet.range("O2").value if sheet.range("O2").value else "0",
        "ang_gb": sheet.range("O4").value if sheet.range("O4").value else "0",
        "uebern_lfm": sheet.range("O3").value if sheet.range("O3").value else "0",
        "uebern_gb": sheet.range("O5").value if sheet.range("O5").value else "0"
    }

    table_data = []
    declined_table_data = []
    row = 2
    while sheet.range(f"A{row}").value: 
        inhalt = sheet.range(f"A{row}").value or ''
        zeitraum = sheet.range(f"B{row}").value or ''
        lfm = sheet.range(f"C{row}").value or '0'
        gb = sheet.range(f"D{row}").value or '0'
        medium = sheet.range(f"E{row}").value or ''
        bewertung = sheet.range(f"G{row}").value or ''
        begruendung = sheet.range(f"H{row}").value or ''
        begruendung_kommentar = sheet.range(f"I{row}").value or ''
        ueber_lfm = sheet.range(f"J{row}").value or '0'
        ueber_gb = sheet.range(f"K{row}").value or '0'
        hasBoth =  True if gb != '0' and lfm != '0' else False

        aktengruppe = f"{inhalt}, {zeitraum},"
        if hasBoth:
            aktengruppe += f" {lfm} Lfm, {gb} GB"
        elif lfm != '0':
            aktengruppe += f" {lfm} Lfm"
        else:
            aktengruppe += f" {gb} GB"
        aktengruppe += f" ({medium})"

        vereinbarung = f"{bewertung} \nBegründung: {begruendung} \n{begruendung_kommentar}"
        if hasBoth:
            uebernommene_menge = f'{ueber_lfm}/{ueber_gb}'
            declined_amount = f'{lfm}/{gb}'
        elif ueber_lfm != '0':
            uebernommene_menge = ueber_lfm
            declined_amount = lfm
        else:
            uebernommene_menge = ueber_gb
            declined_amount = gb

        if(begruendung == "Aufbewahrungsfrist noch laufend"):
            declined_table_data.append([aktengruppe, declined_amount])
        else:
            table_data.append([aktengruppe, vereinbarung, uebernommene_menge])

        row += 1

    return data, table_data, declined_table_data


def replacePlaceholdersInDoc(word, data: dict, contact_sheet) -> None:
    try:
        header_range = contact_sheet.range("A1:Z1")
        stazh_name = "<StAZHName>"
        stazh_mail = "<StAZHMail>"
        stazh_nummer = "<StAZHNummer>"

        for cell in header_range:
            if data.get("kuerz").lower() in str(cell.value).lower():
                column_index = cell.column
                stazh_name = contact_sheet.cells(2, column_index).value or "<StAZHName>"
                stazh_mail = contact_sheet.cells(3, column_index).value or "<StAZHMail>"
                stazh_nummer = contact_sheet.cells(4, column_index).value or "<StAZHNummer>"
                break

        replacements = {
            "<StAZHKürz>": data.get("kuerz"),
            "<ErstellungsDatum>": datetime.now().strftime("%d.%m.%Y"),
            "<AmtZeichen>": data.get("amt_zeichen"),
            "<AmtPName>": data.get("amt_pname"),
            "<AmtMail>": data.get("amt_mail"),
            "<BesDatum>": datetime.strftime(data.get("bes_datum"), "%d.%m.%Y"),
            "<AblJahr>": datetime.now().year,
            "<AngDatum>": datetime.strptime(data.get("ang_datum"), "%d.%m.%Y"),
            "<AngLfm>": data.get("ang_lfm"),
            "<AngGB>": data.get("ang_gb"),
            "<ÜbernLfm>": data.get("uebern_lfm"),
            "<ÜbernGB>": data.get("uebern_gb"),
            "<AmtName>": data.get("amt_name"),
            "<ÜbernDatum>": datetime.strftime(data.get("uebern_datum"), "%d.%m.%Y"),
            "<DirName>": data.get("dir_name"),
            "<StAZHName>": stazh_name,
            "<StAZHNummer>": stazh_mail,
            "<StAZHMail>": stazh_nummer
        }

        for placeholder, value in replacements.items():
            word.Selection.Find.Execute(placeholder, False, False, False, False, False, True, 1, False, value, 2)
            
        for section in word.ActiveDocument.Sections:
            for footer in section.Footers:
                footer.Range.find.Execute("<ErstellungsDatum>", False, False, False, False, False, True, 1, False, replacements["<ErstellungsDatum>"], 2)
                footer.Range.find.Execute("<AmtZeichen>", False, False, False, False, False, True, 1, False, replacements["<AmtZeichen>"], 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter in Word: {e}")


def insertTableDataIntoWord(word, table_data, declined_table_data):
    try:
        doc = word.ActiveDocument
        acceptedFiles = doc.Tables[2]
        declinedFiles = doc.Tables[3]
        declinedText = "Die Aufbewahrungsfrist der folgenden Aktengruppen ist noch nicht abgelaufen. Diese müssen weiterhin aufbewahrt und bei Ablauf der Aufbewahrungsfrist erneut dem StAZH angeboten werden:"

        for row in range(acceptedFiles.Rows.Count, 1, -1):
            acceptedFiles.Rows(row).Delete()

        for i, row in enumerate(table_data):
            acceptedFiles.Rows.Add()
            if re.search(r',\s*(\d{4})', row[0]):
                row[0] = re.sub(r',\s*(\d{4})', r', \n\1', row[0])
            acceptedFiles.Cell(i + 2, 1).Range.Text = row[0]
            acceptedFiles.Cell(i + 2, 2).Range.Text = row[1]
            acceptedFiles.Cell(i + 2, 3).Range.Text = row[2] 

        if len(declined_table_data) > 0:
            for row in range(declinedFiles.Rows.Count, 1, -1):
                declinedFiles.Rows(row).Delete()

            for i, row in enumerate(declined_table_data):
                declinedFiles.Rows.Add()
                if re.search(r',\s*(\d{4})', row[0]):
                    row[0] = re.sub(r',\s*(\d{4})', r', \n\1', row[0])
                declinedFiles.Cell(i + 2, 1).Range.Text = row[0]
                declinedFiles.Cell(i + 2, 2).Range.Text = row[1]
        else:
            declinedFiles.Delete()
            word.Selection.Find.Execute(declinedText, False, False, False, False, False, True, 1, False, "", 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Einfügen der Tabellenwerte in Word: {e}")
        print(f"Fehler beim Einfügen der Tabellenwerte in Word: {e}")