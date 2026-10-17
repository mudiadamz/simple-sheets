"""App name / version, shared by the app, the installer and the build. Also tells simplecore who the app is
(imported before any simplecore module opens a settings file: simple_sheets.py and tests/conftest.py import it
first)."""
import os

from simplecore import appinfo, google_drive

APP_NAME = "Simple Sheets"
APP_ID = "SimpleSheets"
APP_VERSION = "2.0.0"
EXE_NAME = "SimpleSheets.exe"

appinfo.configure(
    name=APP_NAME, id=APP_ID, version=APP_VERSION, exe=EXE_NAME,
    home_env="SIMPLE_SHEETS_HOME",               # tests point the settings folder elsewhere with this
    repo="https://github.com/mudiadamz/simple-sheets",
    description="A light viewer / editor for spreadsheet files (.xlsx, .xlsm, .xlsb, .xls, .ods, .csv).",
    extensions=(".xlsx", ".xlsm", ".xlsb", ".xls", ".ods", ".csv", ".tsv"),
    file_kind="Spreadsheet", doc_noun="spreadsheet", icon="simple_sheets.ico",
    root=os.path.dirname(os.path.abspath(__file__)),
)

XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
SHEET = "application/vnd.google-apps.spreadsheet"
google_drive.setup(SHEET, XLSX, ".xlsx",
                   {".xlsx": XLSX, ".xlsm": "application/vnd.ms-excel.sheet.macroEnabled.12",
                    ".xls": "application/vnd.ms-excel", ".csv": "text/csv", ".tsv": "text/tab-separated-values",
                    ".ods": "application/vnd.oasis.opendocument.spreadsheet"}, "Google Sheet")
