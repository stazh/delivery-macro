# Ablieferungsmakro Setup Guide

This guide explains how to set up and run the **Ablieferungsmakro** Python application.

---

## System Requirements

- Python **3.11.x** (64-bit)  
- Windows operating system  
- Microsoft Excel installed (required for `xlwings`)  
- Microsoft Word installed (required for `pywin32` Word automation)
- Python library `openpyxl` (for Excel file handling)

---

## Setup Instructions

Follow these steps to set up the environment and run the application:

1. Open a terminal or PowerShell in the `Skripts` folder of the project.
2. Create a virtual environment named `.venv`:

    ```powershell
    python -m venv .venv
    ```

3. Activate the virtual environment:

    ```powershell
    ./.venv/Scripts/Activate.ps1
    ```

4. Install required packages:

    ```powershell
    python -m pip install pywin32 xlwings openpyxl
    ```

5. (Optional) If running on DAP or another system, edit `.venv\pyvenv.cfg` and adjust the paths to point to the correct Python installation. This ensures the virtual environment works correctly in that system.

6. Now you can launch the **Ablieferungsmakro** and all functions should be available.

---

## Notes

- Always ensure Excel and Word are installed and accessible.  
- If file paths change (e.g., moving the script or data files), update the paths in `pyvenv.cfg` in the .venv folder.  
- The application interacts with Excel sheets and Word documents, so make sure all required files exist at the specified paths.

---

## Quick Start Summary

1. Open terminal in `Skripts` folder  
2. `python -m venv .venv`  
3. `./.venv/Scripts/Activate.ps1`   
4. `python -m pip install pywin32 xlwings`  
5. (Optional) Adjust `.venv\pyvenv.cfg` if running on DAP or another system  
6. Run the Ablieferungsmakro.xlsm  


## Folderstructure

Project Root
│
├── Ablieferungsmakro.xlsm
├── Skripts
│   ├── .venv
│   ├── main.pyz
│   └── config.py
├── Daten
│   └── Stammdaten.xlsx
├── Glossar
│   └── Glossar.xlsx
├── Generierte Dokumente
│   └── (empty folder for generated documents)
├── Vorlagen
    ├── Ablieferungsvereinbarung.docx
    ├── Ablieferungsverzeichnis.xlsx
    ├── Aktenangebotsformular.docx
    └── Angebotstabelle.xlsx