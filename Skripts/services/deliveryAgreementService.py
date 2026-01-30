import sys
import os
from tkinter import messagebox
from datetime import datetime
import xlwings as xw
from docx import Document
import shutil  # Zum Kopieren der Vorlage

def createDeliveryAgreement() -> None:
    try:
        # Excel-Instanz starten und die aktive Arbeitsmappe laden
        app = xw.apps.active
        app.visible = False
        workbook = app.books['Ablieferungsmakro.xlsm']

        # Stammdaten-Datei laden, um den Pfad zur Word-Vorlage zu holen
        stammdaten_workbook = xw.Book('./Daten/Stammdaten.xlsx')
        dateipfade_sheet = stammdaten_workbook.sheets['Dateipfade']

        # Relativen Pfad zur Vorlage holen
        template_path = dateipfade_sheet.range('B1').value
        stammdaten_workbook.close()

        # Sicherstellen, dass der Pfad korrekt ist
        if not template_path or not os.path.exists(template_path):
            messagebox.showerror("Fehler", f"Der angegebene Pfad zur Vorlage '{template_path}' ist ungültig.")
            return

        # Ausgabe des Pfades für Debugging
        print(f"Vorlage wird geöffnet: {template_path}")

        # Daten aus der Excel-Datei extrahieren
        data = importDeliveryAgreementExcel(workbook)

        # Erstellen einer Kopie der Vorlage als .docx
        copy_path = f"Ablieferungsvereinbarung_{datetime.now().strftime('%d.%m.%Y')}.docx"
        shutil.copy(template_path, copy_path)
        print(f"Vorlage kopiert nach: {copy_path}")

        # Dokument mit python-docx öffnen
        doc = Document(copy_path)
        print("Dokument wurde erfolgreich geöffnet.")

        # Platzhalter ersetzen
        replacePlaceholdersInDoc(doc, data)

        # Speichern der neuen Datei
        output_path = f'Ablieferungsvereinbarung_{datetime.now().strftime("%d.%m.%Y")}.docx'
        doc.save(output_path)
        print(f"Dokument wurde gespeichert unter: {output_path}")

        # Erfolgsmeldung
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
        # Ersetzen der Platzhalter mit den tatsächlichen Werten
        replacements = {
            "<StAZHKürz>": data["kuerz"],
            "<StAZHName>": data["amt_name"],
            "<StAZHNummer>": data["amt_zeichen"],
            "<StAZHMail>": data["amt_mail"],
            "<ErstellungsDatum>": datetime.now().strftime("%d.%m.%Y"),
            "<AmtZeichen>": data["amt_zeichen"],
            "<AmtPName>": data["amt_pname"],
            "<AmtMail>": data["amt_mail"],
            "<BesDatum>": data["bes_datum"],
            "<AblJahr>": datetime.now().year,
            "<AngDatum>": data["ang_datum"],
            "<AngLfm>": data["ang_lfm"],
            "<AngGB>": data["ang_gb"],
            "<ÜbernLfm>": data["uebern_lfm"],
            "<ÜbernGB>": data["uebern_gb"],
            "<AmtName>": data["amt_name"],
            "<ÜbernDatum>": data["uebern_datum"],
            "<DirName>": data["dir_name"]
        }

        # Durch alle Absätze gehen und Platzhalter ersetzen
        for paragraph in doc.paragraphs:
            for placeholder, value in replacements.items():
                if placeholder in paragraph.text:
                    # Platzhalter ersetzen
                    inline = paragraph.runs
                    for run in inline:
                        if placeholder in run.text:
                            run.text = run.text.replace(placeholder, value)

        # Durch alle Tabellen gehen und Platzhalter ersetzen
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for placeholder, value in replacements.items():
                        if placeholder in cell.text:
                            cell.text = cell.text.replace(placeholder, value)

    except Exception as e:
        messagebox.showerror("Fehler", f"Fehler beim Ersetzen der Platzhalter in Word: {e}")
        print(f"Fehler beim Ersetzen der Platzhalter in Word: {e}")

