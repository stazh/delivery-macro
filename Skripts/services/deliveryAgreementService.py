import os
import re
from tkinter import messagebox
from datetime import datetime
import shutil
import config

# Versuche, xlwings zu importieren
try:
    import xlwings as xw
except ImportError as e:
    messagebox.showerror("Fehler", f"Fehler beim Importieren von xlwings: {e}")
    raise

# Versuche, win32com.client zu importieren
try:
    import win32com.client as win32
except ImportError as e:
    messagebox.showerror("Fehler", f"Fehler beim Importieren von win32com.client: {e}")
    raise

def create_delivery_agreement() -> None:
    """Erstellt eine Ablieferungsvereinbarung basierend auf den Excel-Daten und einer Word-Vorlage."""
    try:
        # Excel- und Word-Initialisierung
        app = xw.apps.active
        app.screen_updating = False
        workbook = app.books[config.DELIVERY_MACRO_FILE_NAME]
        stammdaten_workbook = app.books.open(config.DATA_FILE_PATH)
        
        # Vorlagenpfad auslesen und sicherstellen, dass die Datei existiert
        template_path = stammdaten_workbook.sheets[config.DATA_PATH_SHEET_NAME].range(config.DATA_PATH_FIELD_DELIVERY_AGREEMENT).value
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Ungültiger Vorlagenpfad: {template_path}")
            return

        # Daten importieren
        data, table_data, declined_table_data = import_delivery_data(workbook)

        # Zielordner für das Dokument aus config lesen und sicherstellen, dass der Ordner existiert
        output_folder = config.CREATED_DOCUMENT_FOLDER_PATH
        if not os.path.exists(output_folder):
            os.makedirs(output_folder)

        # Sicherstellen, ob die Datei bereits existiert, bevor die Word-Datei erstellt wird
        document_number = 1
        word_filename = os.path.join(output_folder, f"{document_number:02d}_{config.CREATED_DELIVERY_AGREEMENT_FILE_NAME}")
        while os.path.isfile(word_filename):
            document_number += 1
            word_filename = os.path.join(output_folder, f"{document_number:02d}_{config.CREATED_DELIVERY_AGREEMENT_FILE_NAME}")

        shutil.copy(template_path, word_filename)
        
        if not os.path.exists(word_filename):
            messagebox.showerror("Fehler", f"Die Datei konnte nicht gefunden werden: {word_filename}")
            return

        # Umwandlung des relativen Pfades in einen absoluten Pfad
        word_filename_abs = os.path.abspath(word_filename)
        if not os.path.exists(word_filename_abs):
            messagebox.showerror("Fehler", f"Das Ziel-Wortdokument wurde nicht gefunden: {word_filename_abs}")
            return

        # Word-Dokument öffnen und Platzhalter ersetzen
        try:
            word = win32.DispatchEx("Word.Application")
            word.visible = True
            word.Documents.Open(word_filename_abs)

            # Platzhalter ersetzen und Tabellen einfügen
            replace_placeholders_in_doc(word, data, stammdaten_workbook.sheets[config.CONTACT_SHEET_NAME])
            insert_table_data(word, table_data, declined_table_data)

        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Öffnen und Bearbeiten des Word-Dokuments: {e}")
            return
        
        stammdaten_workbook.close()
        word.ActiveDocument.Save()

        messagebox.showinfo("Erfolg", "Die Ablieferungsvereinbarung wurde erfolgreich erstellt.")

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Erstellen der Ablieferungsvereinbarung: {e}")


def import_delivery_data(workbook) -> dict:
    """Importiert relevante Daten aus dem aktuellen Angebot (Excel)."""
    sheet = workbook.sheets[config.CURRENT_OFFER_SHEET_NAME]
    
    # Wichtige Felder aus dem Excel-Sheet extrahieren
    data = {
        "kuerz": sheet.range(config.CURRENT_OFFER_FIELD_KUERZ).value or config.KUERZ_PLACEHOLDER,
        "amt_zeichen": sheet.range(config.CURRENT_OFFER_FIELD_AMT_ZEICHEN).value or config.AMT_ZEICHEN_PLACEHOLDER,
        "amt_pname": sheet.range(config.CURRENT_OFFER_FIELD_AMT_P_NAME).value or config.AMT_P_NAME_PLACEHOLDER,
        "bes_datum": sheet.range(config.CURRENT_OFFER_FIELD_BES_DATUM).value or config.BES_DATUM_PLACEHOLDER,
        "ang_datum": sheet.range(config.CURRENT_OFFER_FIELD_ANG_DATUM).value or config.ANG_DATUM_PLACEHOLDER,
        "uebern_datum": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_DATUM).value or config.UEBERN_DATUM_PLACEHOLDER,
        "dir_name": sheet.range(config.CURRENT_OFFER_FIELD_DIR_NAME).value or config.DIR_NAME_PLACEHOLDER,
        "ang_lfm": sheet.range(config.CURRENT_OFFER_FIELD_ANG_LFM).value or config.DEFAULT_VALUE_LFM_GB,
        "ang_gb": sheet.range(config.CURRENT_OFFER_FIELD_ANG_GB).value or config.DEFAULT_VALUE_LFM_GB,
        "uebern_lfm": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_LFM).value or config.DEFAULT_VALUE_LFM_GB,
        "uebern_gb": sheet.range(config.CURRENT_OFFER_FIELD_UEBERN_GB).value or config.DEFAULT_VALUE_LFM_GB
    }

    table_data, declined_table_data = [], []
    row = 2
    while sheet.range(f"A{row}").value:  # Durchläuft alle Zeilen
        # Jede Zeile als Dictionary abspeichern
        row_data = {
            "inhalt": sheet.range(f"A{row}").value or '',
            "zeitraum": sheet.range(f"B{row}").value or '',
            "lfm": sheet.range(f"C{row}").value or config.DEFAULT_VALUE_LFM_GB,
            "gb": sheet.range(f"D{row}").value or config.DEFAULT_VALUE_LFM_GB,
            "medium": sheet.range(f"E{row}").value or '',
            "ueber_medium": sheet.range(f"L{row}").value or '',
            "bewertung": sheet.range(f"G{row}").value or '',
            "begruendung": sheet.range(f"H{row}").value or '',
            "begruendung_kommentar": sheet.range(f"I{row}").value or '',
            "ueber_lfm": sheet.range(f"J{row}").value or config.DEFAULT_VALUE_LFM_GB,
            "ueber_gb": sheet.range(f"K{row}").value or config.DEFAULT_VALUE_LFM_GB
        }
        # Aktengruppen erstellen
        act_group = f"{row_data['inhalt']},{row_data['zeitraum']},"
        declined_act_group = act_group
        if f"{row_data['lfm']}" != config.DEFAULT_VALUE_LFM_GB and f"{row_data['gb']}" != config.DEFAULT_VALUE_LFM_GB:
            act_group += f" {row_data['lfm']} Lfm, {row_data['gb']} GB"
        else:
            lfm_or_gb = f"{row_data['lfm']} Lfm" if f"{row_data['lfm']}" != config.DEFAULT_VALUE_LFM_GB else f"{row_data['gb']} GB"
            declined_amount = lfm_or_gb
            act_group += f" {lfm_or_gb}"
        act_group += f" ({row_data['medium']})"
        
        vereinbarung = f"{row_data['bewertung']}\nBegründung: {row_data['begruendung']}\n{row_data['begruendung_kommentar']}"

        # Je nach Bewertung in die entsprechende Tabelle einfügen
        if row_data['begruendung'] == config.WORD_SHOW_DECLINED_LIST:
            declined_amount = ""
            if f"{row_data['lfm']}" != config.DEFAULT_VALUE_LFM_GB and f"{row_data['gb']}" != config.DEFAULT_VALUE_LFM_GB:
                declined_amount = f"{row_data['lfm']}/{row_data['gb']}"
            else:
                declined_amount = f"{row_data['lfm']} Lfm" if f"{row_data['lfm']}" != config.DEFAULT_VALUE_LFM_GB else f"{row_data['gb']} GB"

            medium = f" ({row_data['medium']})" if row_data['medium'] and row_data['medium'].strip() != "" else ""
            declined_table_data.append([declined_act_group.removesuffix(","), f"{declined_amount}{medium}"])
        else:
            accepted_amount = ""
            if f"{row_data['ueber_lfm']}" != config.DEFAULT_VALUE_LFM_GB and f"{row_data['ueber_gb']}" != config.DEFAULT_VALUE_LFM_GB:
                accepted_amount = f"{row_data['ueber_lfm']}/{row_data['ueber_gb']}"
            else:
                if f"{row_data['ueber_lfm']}" != config.DEFAULT_VALUE_LFM_GB:
                    accepted_amount = f"{row_data['ueber_lfm']} Lfm"
                elif f"{row_data['ueber_gb']}" != config.DEFAULT_VALUE_LFM_GB:
                    accepted_amount = f"{row_data['ueber_gb']} GB"
                else:
                    accepted_amount = config.DEFAULT_VALUE_LFM_GB

            medium = f" ({row_data['ueber_medium']})" if row_data['ueber_medium'] and row_data['ueber_medium'].strip() != "" else ""
            table_data.append([act_group, vereinbarung, f"{accepted_amount}{medium}"])

        row += 1

    return data, table_data, declined_table_data


def replace_placeholders_in_doc(word, data: dict, contact_sheet) -> None:
    """Ersetzt die Platzhalter im Word-Dokument mit den tatsächlichen Werten."""
    try:
        header_range = contact_sheet.range(config.DATA_CONTACT_RANGE)
        stazh_name, stazh_mail, stazh_nummer = config.STAZH_NAME_PLACEHOLDER, config.STAZH_MAIL_PLACEHOLDER, config.STAZH_NUMMER_PLACEHOLDER
        
        # Findet den entsprechenden Kontakt in der Tabelle und aktualisiert die Werte
        found = False
        for cell in header_range:
            if data["kuerz"].lower() in str(cell.value).lower():
                column_index = cell.column

                stazh_name = contact_sheet.cells(2, column_index).value or config.STAZH_NAME_PLACEHOLDER
                stazh_mail = contact_sheet.cells(3, column_index).value or config.STAZH_MAIL_PLACEHOLDER
                stazh_nummer = contact_sheet.cells(4, column_index).value or config.STAZH_NUMMER_PLACEHOLDER

                found = True
                break

        # Wenn nichts gefunden wurde -> Abbruch
        if not found:
            messagebox.showerror(
                "Fehler",
                f"Kein Kontakt für '{data['kuerz']}' gefunden."
            )
            return

        replacements = {
            config.KUERZ_PLACEHOLDER: data["kuerz"],
            config.ERSTELLUNGS_DATUM_PLACEHOLDER: datetime.now().strftime("%d.%m.%Y"),
            config.AMT_ZEICHEN_PLACEHOLDER: data["amt_zeichen"],
            config.AMT_P_NAME_PLACEHOLDER: data["amt_pname"],
            config.BES_DATUM_PLACEHOLDER: data["bes_datum"],
            config.ABL_JAHR_PLACEHOLDER: datetime.now().year,
            config.ANG_DATUM_PLACEHOLDER: data["ang_datum"],
            config.ANG_LFM_PLACEHOLDER: data["ang_lfm"],
            config.ANG_GB_PLACEHOLDER: data["ang_gb"],
            config.UEBERN_LFM_PLACEHOLDER: data["uebern_lfm"],
            config.UEBERN_GB_PLACEHOLDER: data["uebern_gb"],
            config.UEBERN_DATUM_PLACEHOLDER: data["uebern_datum"],
            config.DIR_NAME_PLACEHOLDER: data["dir_name"],
            config.STAZH_NAME_PLACEHOLDER: stazh_name,
            config.STAZH_MAIL_PLACEHOLDER: stazh_mail,
            config.STAZH_NUMMER_PLACEHOLDER: stazh_nummer
        }

        # Ersetzen der Platzhalter
        for placeholder, value in replacements.items():
            word.Selection.Find.Execute(placeholder, False, False, False, False, False, True, 1, False, value, 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter in Word: {e}")


def insert_table_data(word, table_data, declined_table_data) -> None:
    """Fügt die Tabellenwerte in das Word-Dokument ein."""
    try:
        # Zugriff auf das aktive Word-Dokument und Tabellen
        doc = word.ActiveDocument
        acceptedFiles = doc.Tables[2]
        declinedFiles = doc.Tables[3]
        declinedText = config.WORD_DECLINED_TEXT

        # Entfernt alle bestehenden Zeilen in der Tabelle für akzeptierte Dateien
        for row in range(acceptedFiles.Rows.Count, 1, -1):
            acceptedFiles.Rows(row).Delete()

        # Fügt neue Zeilen in die Tabelle für akzeptierte Dateien ein
        for i, row in enumerate(table_data):
            acceptedFiles.Rows.Add()
            # Bearbeitet den Text in der ersten Spalte, wenn er mit dem Muster übereinstimmt
            if re.search(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, row[0]):
                row[0] = re.sub(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, r',\n\1', row[0])
            acceptedFiles.Cell(i + 2, 1).Range.Text = row[0]
            acceptedFiles.Cell(i + 2, 2).Range.Text = row[1]
            acceptedFiles.Cell(i + 2, 3).Range.Text = row[2]

        # Wenn es abgelehnte Dateien gibt, füge sie in die Tabelle für abgelehnte Dateien ein
        if len(declined_table_data) > 0:
            for row in range(declinedFiles.Rows.Count, 1, -1):
                declinedFiles.Rows(row).Delete()

            # Fügt abgelehnte Dateien in die Tabelle ein
            for i, row in enumerate(declined_table_data):
                declinedFiles.Rows.Add()
                if re.search(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, row[0]):
                    row[0] = re.sub(config.REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES, r',\n\1', row[0])
                declinedFiles.Cell(i + 2, 1).Range.Text = row[0]
                declinedFiles.Cell(i + 2, 2).Range.Text = row[1]
        else:
            declinedFiles.Delete()
            word.Selection.Find.Execute(declinedText, False, False, False, False, False, True, 1, False, "", 2)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Einfügen der Tabellenwerte in Word: {e}")
