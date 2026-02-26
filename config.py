# ------------------------
# 1. Dateinamen und Pfade
# ------------------------
DELIVERY_MACRO_FILE_NAME = "Ablieferungsmakro.xlsm"
DATA_FILE_PATH = "./Daten/Stammdaten.xlsx"
CREATED_DELIVERY_AGREEMENT_FILE_NAME = "Ablieferungsvereinbarung.docx"
CREATED_DOCUMENT_FOLDER_PATH = "./Erstellte Dokumente"
DELIVERY_LIST_FILE_NAME = "Ablieferungsverzeichnis.xlsx"

# --------------------------
# 2. Tabellenblattnamen
# --------------------------
DATA_PATH_SHEET_NAME = "Dateipfade"
CONTACT_SHEET_NAME = "Kontakte"
CURRENT_OFFER_SHEET_NAME = "Angebot_aktuell"
SELECTION_FIELDS_SHEET_NAME = "Auswahlfelder"
DELIVERY_LIST_SHEET_NAME = "Ablieferungsverzeichnis"

# -----------------------------------------
# 3. Aktuelles Angebot: Auswahl
# -----------------------------------------
SELECTION_REASON_OFFER_LENGTH = "G2:G6268"
SELECTION_REASON_OFFER_RANGE = "A1:D1"
SELECTION_SECTION = "AZ"
SELECTION_COMPLETE_TAKEOVER_FIELDS = "A2:A11"
SELECTION_PARTIAL_TAKEOVER_FIELDS = "B2:B5"
SELECTION_NO_TAKEOVER_FIELDS = "C2:C9"
SELECTION_CANT_BE_OFFERED_FIELDS = "D2"

# -----------------------------
# 4. Zellenbezüge
# -----------------------------
DATA_PATH_FIELD_DELIVERY_AGREEMENT = "B1"
DATA_PATH_FIELD_DELIVERY_LIST = "B2"
DATA_PATH_FIELD_OFFER_TABLE = "B4"
CURRENT_OFFER_FIELD_KUERZ = "O9"
CURRENT_OFFER_FIELD_AMT_MAIL = "Z6"
CURRENT_OFFER_FIELD_AMT_ZEICHEN = "O13"
CURRENT_OFFER_FIELD_AMT_P_NAME = "O8"
CURRENT_OFFER_FIELD_AMT_NAME = "O7"
CURRENT_OFFER_FIELD_BES_DATUM = "O10"
CURRENT_OFFER_FIELD_ANG_DATUM = "O11"
CURRENT_OFFER_FIELD_UEBERN_DATUM = "O12"
CURRENT_OFFER_FIELD_DIR_NAME = "O6"
CURRENT_OFFER_FIELD_ANG_LFM = "O2"
CURRENT_OFFER_FIELD_ANG_GB = "O4"
CURRENT_OFFER_FIELD_UEBERN_GB = "O5"
CURRENT_OFFER_FIELD_UEBERN_LFM = "O3"
DELIVERY_LIST_PLACEHOLDER_RANGE = "A1:G20"
DATA_CONTACT_RANGE = "A1:Z1"

# ----------------------
# 5. Platzhalter
# ----------------------
KUERZ_PLACEHOLDER = "<StAZHKürz>"
AMT_MAIL_PLACEHOLDER = "<AmtMail>"
AMT_ZEICHEN_PLACEHOLDER = "<AmtZeichen>"
AMT_P_NAME_PLACEHOLDER = "<AmtPName>"
AMT_NAME_PLACEHOLDER = "<AmtName>"
BES_DATUM_PLACEHOLDER = "<BesDatum>"
ANG_DATUM_PLACEHOLDER = "<AngDatum>"
UEBERN_DATUM_PLACEHOLDER = "<ÜbernDatum>"
DIR_NAME_PLACEHOLDER = "<DirName>"
STAZH_NAME_PLACEHOLDER = "<StAZHName>"
STAZH_MAIL_PLACEHOLDER = "<StAZHMail>"
STAZH_NUMMER_PLACEHOLDER = "<StAZHNummer>"
ERSTELLUNGS_DATUM_PLACEHOLDER = "<ErstellungsDatum>"
ABL_JAHR_PLACEHOLDER = "<AblJahr>"
ANG_GB_PLACEHOLDER = "<AngGB>"
ANG_LFM_PLACEHOLDER = "<AngLfm>"
UEBERN_LFM_PLACEHOLDER = "<ÜbernLfm>"
UEBERN_GB_PLACEHOLDER = "<ÜbernGB>"
ABL_NUMMER_PLACEHOLDER = "<AblNummer>"

# -------------------------
# 6. Standardwerte
# -------------------------
DEFAULT_VALUE_LFM_GB = "0"

# ------------------------
# 7. Reguläre Ausdrücke
# ------------------------
REGEX_WORD_DELIVERY_AGREEMENT_ACCEPTED_DECLINED_FILES = ",\\s*(\\d{4})"
REGEX_WORD_DELIVERY_AGREEMENT_SEARCH_YEAR = ",\s*(\d{4}(?:\s*[-+]\s*\d{4})*)"
REGEX_WORD_DELIVERY_AGREEMENT_SEARCH_MEDIUM = "\(([^)]+)\)(?!.*\()"
REGEX_WORD_DELIVERY_AGREEMENT_SEARCH_GB = "(\d+(\.\d+)?)\s*GB"
REGEX_WORD_DELIVERY_AGREEMENT_SEARCH_LFM = "(\d+(\.\d+)?)\s*Lfm"

# ------------------------
# 8. Nachrichten
# ------------------------
WORD_SHOW_DECLINED_LIST = "Aufbewahrungsfrist noch laufend"
WORD_DECLINED_TEXT = "Die Aufbewahrungsfrist der folgenden Aktengruppen ist noch nicht abgelaufen. Diese müssen weiterhin aufbewahrt und bei Ablauf der Aufbewahrungsfrist erneut dem StAZH angeboten werden:"

# ----------------------------------------------
# 9. Indizes der Tabellen im Liefervereinbarungsdokument
# ----------------------------------------------
WORD_TABLE_INDEX = 2
WORD_FIRST_DATA_ROW = 3
COL_INHALT = 1
COL_ZEITRAUM = 2
COL_LFM = 3
COL_GB = 4
COL_MEDIUM = 5
ROW_BEHOERDE = 6
ROW_AMT = 7 
ROW_ZUSTAENDIG = 8 
ROW_DATUM_ANGEBOT = 11
COL_METADATA = 2
START_DATA_ROW = 2

# ------------------------------
# 10. Texte zur Extraktion im Liefervereinbarungsdokument
# ------------------------------
WORD_DIRECTION_TEXT = "Direktion (Behörde):"
WORD_OFFICE_TEXT = "Amtsstelle:"
WORD_RESPONSIBLE_TEXT = "Zuständig in der Amtsstelle:"
WORD_OFFER_DATE_TEXT = "Datum des Angebots:"