import os
import re
from tkinter import messagebox
from datetime import datetime
import xlwings as xw
import shutil
import win32com.client as win32
import config

def createDeliveryAgreement() -> None: 
    try: 
        app = xw.apps.active 
        app.visible = False 
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME]
        stammdaten_workbook = xw.Book(config.DATA_FILE_PATH)
        dateipfade_sheet = stammdaten_workbook.sheets[config.DATA_PATH_SHEET_NAME]
        contact_sheet = stammdaten_workbook.sheets[config.CONTACT_SHEET_NAME]
        template_path = dateipfade_sheet.range(config.DATA_PATH_FIELD_DELIVERY_AGREEMENT).value
        if not template_path or not os.path.exists(template_path): 
            messagebox.showerror("Fehler", f"Der angegebene Pfad zur Vorlage '{template_path}' ist ungültig.") 
            return 
        
        data, table_data, declined_table_data = importDeliveryAgreementExcel(workbook)
        
        wordPath = os.path.abspath(config.CREATED_DELIVERY_AGREEMENT_FILE_NAME)
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
        
        messagebox.showinfo("Fertig", f"Die Ablieferungsvereinbarung wurde erfolgreich erstellt")

    except Exception as e: 
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Ablieferungsvereinbarung: {e}") 


def importDeliveryAgreementExcel(workbook) -> dict:
    sheet = workbook.sheets[config.CURRENT_OFFER_SHEET_NAME]
    
    data = {
        "kuerz": sheet.range(config.CURRENT_OFFER_FIELD_KUERZ).value if sheet.range(config.CURRENT_OFFER_FIELD_KUERZ).value else config.KUERZ_PLACEHOLDER,
        "amt_mail": sheet.range(config.CURRENT_OFFER_FIELD_AMT_MAIL).value if sheet.range(config.CURRENT_OFFER_FIELD_AMT_MAIL).value else config.AMT_MAIL_PLACEHOLDER,
        "amt_zeichen": sheet.range(config.CURRENT_OFFER_FIELD_AMT_ZEICHEN).value if sheet.range(config.CURRENT_OFFER_FIELD_AMT_ZEICHEN).value else config.AMT_ZEICHEN_PLACEHOLDER,
        "amt_pname": sheet.range(config.CURRENT_OFFER_FIELD_AMT_P_NAME).value if sheet.range(config.CURRENT_OFFER_FIELD_AMT_P_NAME).value else config.AMT_P_NAME_PLACEHOLDER,
        "amt_name": sheet.range(config.CURRENT_OFFER_FIELD_AMT_NAME).value if sheet.range(config.CURRENT_OFFER_FIELD_AMT_NAME).value else config.AMT_NAME_PLACEHOLDER,
        "bes_datum": sheet.range(config.CURRENT_OFFER_FIELD_BES_DATUM).value if sheet.range(config.CURRENT_OFFER_FIELD_BES_DATUM).value else config.BES_DATUM_PLACEHOLDER,
        "ang_datum": sheet.range(config.CURRENT_OFFER_FIELD_ANG_DATUM).value if sheet.range(config.CURRENT_OFFER_FIELD_ANG_DATUM).value else config.ANG_DATUM_PLACEHOLDER,
        "uebern_datum": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_DATUM).value if sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_DATUM).value else config.UEBERN_DATUM_PLACEHOLDER,
        "dir_name": sheet.range(config.CURRENT_OFFER_FIELD_DIR_NAME).value if sheet.range(config.CURRENT_OFFER_FIELD_DIR_NAME).value else config.DIR_NAME_PLACEHOLDER,
        "ang_lfm": sheet.range(config.CURRENT_OFFER_FIELD_ANG_LFM).value if sheet.range(config.CURRENT_OFFER_FIELD_ANG_LFM).value else config.DEFAULT_VALUE_LFM_GB,
        "ang_gb": sheet.range(config.CURRENT_OFFER_FIELD_ANG_GB).value if sheet.range(config.CURRENT_OFFER_FIELD_ANG_GB).value else config.DEFAULT_VALUE_LFM_GB,
        "uebern_lfm": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_LFM).value if sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_LFM).value else config.DEFAULT_VALUE_LFM_GB,
        "uebern_gb": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_GB).value if sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_GB).value else config.DEFAULT_VALUE_LFM_GB
    }

    table_data = []
    declined_table_data = []
    row = 2
    while sheet.range(f"A{row}").value: 
        inhalt = sheet.range(f"A{row}").value or ''
        zeitraum = sheet.range(f"B{row}").value or ''
        lfm = sheet.range(f"C{row}").value or config.DEFAULT_VALUE_LFM_GB
        gb = sheet.range(f"D{row}").value or config.DEFAULT_VALUE_LFM_GB
        medium = sheet.range(f"E{row}").value or ''
        bewertung = sheet.range(f"G{row}").value or ''
        begruendung = sheet.range(f"H{row}").value or ''
        begruendung_kommentar = sheet.range(f"I{row}").value or ''
        ueber_lfm = sheet.range(f"J{row}").value or config.DEFAULT_VALUE_LFM_GB
        ueber_gb = sheet.range(f"K{row}").value or config.DEFAULT_VALUE_LFM_GB
        hasBoth =  True if gb != config.DEFAULT_VALUE_LFM_GB and lfm != config.DEFAULT_VALUE_LFM_GB else False

        aktengruppe = f"{inhalt}, {zeitraum}," 
        if hasBoth:
            aktengruppe += f" {lfm} Lfm, {gb} GB"
        elif lfm != config.DEFAULT_VALUE_LFM_GB:
            aktengruppe += f" {lfm} Lfm"
        else:
            aktengruppe += f" {gb} GB"
        aktengruppe += f" ({medium})"

        vereinbarung = f"{bewertung} \nBegründung: {begruendung} \n{begruendung_kommentar}"
        if hasBoth:
            uebernommene_menge = f'{ueber_lfm}/{ueber_gb}'
            declined_amount = f'{lfm}/{gb}'
        elif ueber_lfm != config.DEFAULT_VALUE_LFM_GB:
            uebernommene_menge = ueber_lfm
            declined_amount = lfm
        else:
            uebernommene_menge = ueber_gb
            declined_amount = gb

        if(begruendung == config.WORD_SHOW_DECLINED_LIST):
            declined_table_data.append([aktengruppe, declined_amount])
        else:
            table_data.append([aktengruppe, vereinbarung, uebernommene_menge])

        row += 1

    return data, table_data, declined_table_data


def replacePlaceholdersInDoc(word, data: dict, contact_sheet) -> None:
    try:
        header_range = contact_sheet.range(config.DATA_CONTACT_RANGE)
        stazh_name = config.STAZH_NAME_PLACEHOLDER
        stazh_mail = config.STAZH_MAIL_PLACEHOLDER
        stazh_nummer = config.STAZH_NUMMER_PLACEHOLDER

        for cell in header_range:
            if data.get("kuerz").lower() in str(cell.value).lower():
                column_index = cell.column
                stazh_name = contact_sheet.cells(2, column_index).value or config.STAZH_NAME_PLACEHOLDER
                stazh_mail = contact_sheet.cells(3, column_index).value or config.STAZH_MAIL_PLACEHOLDER
                stazh_nummer = contact_sheet.cells(4, column_index).value or config.STAZH_NUMMER_PLACEHOLDER
                break

        replacements = {
            config.KUERZ_PLACEHOLDER: data.get("kuerz"),
            config.ERSTELLUNGS_DATUM_PLACEHOLDER: datetime.now().strftime("%d.%m.%Y"),
            config.AMT_ZEICHEN_PLACEHOLDER: data.get("amt_zeichen"),
            config.AMT_P_NAME_PLACEHOLDER: data.get("amt_pname"),
            config.AMT_MAIL_PLACEHOLDER: data.get("amt_mail"),
            config.BES_DATUM_PLACEHOLDER: datetime.strftime(data.get("bes_datum"), "%d.%m.%Y"),
            config.ABL_JAHR_PLACEHOLDER: datetime.now().year,
            config.ANG_DATUM_PLACEHOLDER: datetime.strptime(data.get("ang_datum"), "%d.%m.%Y"),
            config.ANG_LFM_PLACEHOLDER: data.get("ang_lfm"),
            config.ANG_GB_PLACEHOLDER: data.get("ang_gb"),
            config.UEBERN_LFM_PLACEHOLDER: data.get("uebern_lfm"),
            config.UEBERN_GB_PLACEHOLDER: data.get("uebern_gb"),
            config.AMT_NAME_PLACEHOLDER: data.get("amt_name"),
            config.UEBERN_DATUM_PLACEHOLDER: datetime.strftime(data.get("uebern_datum"), "%d.%m.%Y"),
            config.DIR_NAME_PLACEHOLDER: data.get("dir_name"),
            config.STAZH_NAME_PLACEHOLDER: stazh_name,
            config.STAZH_MAIL_PLACEHOLDER: stazh_mail,
            config.STAZH_NUMMER_PLACEHOLDER: stazh_nummer
        }

        for placeholder, value in replacements.items():
            word.Selection.Find.Execute(placeholder, False, False, False, False, False, True, 1, False, value, 2)
            
        for section in word.ActiveDocument.Sections:
            for footer in section.Footers:
                footer.Range.find.Execute(config.ERSTELLUNGS_DATUM_PLACEHOLDER, False, False, False, False, False, True, 1, False, replacements[config.ERSTELLUNGS_DATUM_PLACEHOLDER], 2)
                footer.Range.find.Execute(config.AMT_ZEICHEN_PLACEHOLDER, False, False, False, False, False, True, 1, False, replacements[config.AMT_ZEICHEN_PLACEHOLDER], 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter in Word: {e}")


def insertTableDataIntoWord(word, table_data, declined_table_data):
    try:
        doc = word.ActiveDocument
        acceptedFiles = doc.Tables[2]
        declinedFiles = doc.Tables[3]
        declinedText = config.WORD_DECLINED_TEXT

        for row in range(acceptedFiles.Rows.Count, 1, -1):
            acceptedFiles.Rows(row).Delete()

        for i, row in enumerate(table_data):
            acceptedFiles.Rows.Add()
            if re.search(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, row[0]):
                row[0] = re.sub(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, r', \n\1', row[0])
            acceptedFiles.Cell(i + 2, 1).Range.Text = row[0]
            acceptedFiles.Cell(i + 2, 2).Range.Text = row[1]
            acceptedFiles.Cell(i + 2, 3).Range.Text = row[2] 

        if len(declined_table_data) > 0:
            for row in range(declinedFiles.Rows.Count, 1, -1):
                declinedFiles.Rows(row).Delete()

            for i, row in enumerate(declined_table_data):
                declinedFiles.Rows.Add()
                if re.search(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, row[0]):
                    row[0] = re.sub(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, r', \n\1', row[0])
                declinedFiles.Cell(i + 2, 1).Range.Text = row[0]
                declinedFiles.Cell(i + 2, 2).Range.Text = row[1]
        else:
            declinedFiles.Delete()
            word.Selection.Find.Execute(declinedText, False, False, False, False, False, True, 1, False, "", 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Einfügen der Tabellenwerte in Word: {e}")
        print(f"Fehler beim Einfügen der Tabellenwerte in Word: {e}")
