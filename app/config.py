import os
from pathlib import Path

# Base directory of your project
BASE_DIR = Path(r"C:\Users\AliciaH\OneDrive - buckleycable.com\Documents\Samsara and GPS Validation Check Project\Samsara & GPS Validation")

# Input folders
IMPORT_DIR = BASE_DIR / "Import Files" / "Excel"
PAYCHEX_EXCEL_PATH = IMPORT_DIR / "Paychex_Files"
SAMSARA_EXCEL_PATH = IMPORT_DIR / "Samsara_Files"

# Output folder
EXPORT_DIR = BASE_DIR / "Export Files" / "Excel"
EXPORT_DIR.mkdir(parents=True, exist_ok=True)

# Output file
OUTPUT_EXCEL_PATH = EXPORT_DIR / "TEST.xlsx"

# Flags
VERBOSE = True
EXPORT_OUTPUT = True
