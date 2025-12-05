"""
vba_simulation.py - Automates Excel to run a macro and process data.
Refactored for better error handling, logging, and clarity.
"""

import json
import win32com.client as win32
import os
import logging
import glob


# === SETTINGS ===
# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Utility to find first file by extension in a folder
def find_file_by_ext(folder: str, ext: str) -> str | None:
    files = glob.glob(os.path.join(folder, f"*.{ext}"))
    return files[0] if files else None

EXCEL_PATH = find_file_by_ext(os.path.join(SCRIPT_DIR, "public"), "xlsm")
JSON_OUTPUT_PATH = find_file_by_ext(os.path.join(SCRIPT_DIR, "public"), "json")

INPUT_SHEET = "タイトル"  # Sheet name (can be "Title" or "タイトル")
OUTPUT_SHEET = "タイトル"
MACRO_NAME = "Trimming"

START_ROW = 2  # Excel input starts at row 2
MAX_RECORDS = 8000  # Limit to 10 records for testing

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# =============================
# MAPPING (📌 You must update)
# =============================

# Output Mapping: Excel Column → JSON key in result
OUTPUT_MAPPING = {
    "I": "タイトル",
    "F": "ASIN",
    "G": "巻数",
}

# =============================
def unblock_file(file_path):
    """Remove 'Mark of the Web' from file if it exists"""
    try:
        import subprocess
        cmd = f'powershell -Command "Unblock-File -Path \'{file_path}\'"'
        subprocess.run(cmd, shell=True, capture_output=True)
        logging.info(f"Unblocked file: {file_path}")
    except Exception as e:
        logging.warning(f"Could not unblock file: {e}")


def run_excel_process():
    """
    Automate Excel: extract only OUTPUT_MAPPING columns from the sheet and save output.
    Handles error logging and cleanup.
    """
    unblock_file(os.path.abspath(EXCEL_PATH))
    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = True
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = True
    excel.EnableEvents = True
    wb = None
    try:
        try:
            excel.AutomationSecurity = 1
        except Exception:
            logging.warning("Could not set AutomationSecurity (may require admin rights)")
        excel_file_path = os.path.abspath(EXCEL_PATH)
        logging.info(f"Opening Excel file: {excel_file_path}")
        wb = excel.Workbooks.Open(
            excel_file_path,
            UpdateLinks=0,
            ReadOnly=False,
            Format=None,
            CorruptLoad=0
        )
        logging.info("Workbook opened successfully")
        try:
            out_sheet = wb.Sheets(OUTPUT_SHEET)
        except Exception:
            alt_sheet = "Title" if OUTPUT_SHEET == "タイトル" else "タイトル"
            try:
                out_sheet = wb.Sheets(alt_sheet)
                logging.info(f"Using sheet '{alt_sheet}' instead of '{OUTPUT_SHEET}'")
            except Exception:
                logging.error(f"Could not find sheet '{OUTPUT_SHEET}' or '{alt_sheet}'")
                raise RuntimeError(f"Could not find sheet '{OUTPUT_SHEET}' or '{alt_sheet}'")
        results = []
        row = START_ROW
        while True:
            record = {}
            empty_row = True
            for col, json_key in OUTPUT_MAPPING.items():
                cell = out_sheet.Range(f"{col}{row}")
                value = cell.Value
                record[json_key] = value
                if value not in (None, ""):
                    empty_row = False
                # --- Title color check ---
                if json_key == "タイトル":
                    color_value = cell.DisplayFormat.Interior.Color
                    record["color"] = (color_value == 9895780.0)
                # --- Volume number (巻数) ---
                if json_key == "巻数":
                    if value is None:
                        record["巻数"] = 1
                    else:
                        try:
                            record["巻数"] = int(value)
                        except:
                            record["巻数"] = 1
            if empty_row:
                break
            results.append(record)
            row += 1
        with open(JSON_OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        logging.info(f"Extraction complete! → {JSON_OUTPUT_PATH}")
        return results
    except Exception as e:
        logging.error(f"Error during Excel extraction: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        try:
            excel.Quit()
        except Exception:
            pass

if __name__ == "__main__":
    run_excel_process()
