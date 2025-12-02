
"""
vba_simulation.py - Automates Excel to run a macro and process data.
Refactored for better error handling, logging, and clarity.
"""

import json
import win32com.client as win32
import time
import os
import threading
import win32gui
import win32con
import logging


# === SETTINGS ===
# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(SCRIPT_DIR, "runMacro", "runMacro.xlsm")
JSON_INPUT_PATH = os.path.join(SCRIPT_DIR, "runMacro", "target_macro_input.json")
JSON_OUTPUT_PATH = os.path.join(SCRIPT_DIR, "runMacro", "target_macro_output.json")

INPUT_SHEET = "タイトル"  # Sheet name (can be "Title" or "タイトル")
OUTPUT_SHEET = "タイトル"
MACRO_NAME = "Trimming"

START_ROW = 2  # Excel input starts at row 2
MAX_RECORDS = 10  # Limit to 10 records for testing

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# =============================
# MAPPING (📌 You must update)
# =============================

# Input Mapping: JSON key → Excel Column
# Maps: ID → A, AncestorID → B, AdministrativeTitle → C, AmazonTitle → D
INPUT_MAPPING = {
    "ID": "A",                    # Column A: ID
    "先祖ID": "B",                 # Column B: AncestorID (先祖ID in Japanese)
    "管理タイトル": "C",            # Column C: AdministrativeTitle (管理タイトル in Japanese)
    "Amazonタイトル": "D",          # Column D: AmazonTitle (Amazonタイトル in Japanese)
    "先祖-ASIN": "F",
}

# Output Mapping: Excel Column → JSON key in result
OUTPUT_MAPPING = {
    "I": "タイトル",
    "F": "ASIN",
    "G": "巻数",
}

# =============================

def close_msgbox(window_title_contains="マクロ実行確認", button_name="はい", timeout=15):
    """
    Automatically click button in VBA dialog box.
    Runs in a background thread to monitor for the dialog while macro executes.
    
    Args:
        window_title_contains: Substring to match in dialog window title
        button_name: Text on the button to click (e.g., "はい" for Yes)
        timeout: Maximum time to wait for dialog (seconds)
    
    Returns:
        Thread object that is monitoring for the dialog
    """
    def monitor_and_click():
        """Background thread function that monitors for dialog and clicks button"""
        start_time = time.time()
        dialog_handled = False
        while time.time() - start_time < timeout and not dialog_handled:
            windows = []
            def enum_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if window_title_contains in title or "マクロ" in title or "実行確認" in title:
                            windows.append((hwnd, title))
                    except Exception:
                        pass
                return True
            win32gui.EnumWindows(enum_callback, windows)
            for hwnd, title in windows:
                if not dialog_handled:
                    logging.info(f"Found dialog: {title}")
                    dialog_handled = True
                    try:
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                        time.sleep(0.2)
                        button_clicked = False
                        def find_button(child_hwnd, param):
                            nonlocal button_clicked
                            try:
                                child_text = win32gui.GetWindowText(child_hwnd)
                                if button_name in child_text:
                                    win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                                    logging.info(f"✓ Clicked button: {child_text}")
                                    button_clicked = True
                                    return False
                            except Exception:
                                pass
                            return True
                        win32gui.EnumChildWindows(hwnd, find_button, None)
                        if not button_clicked:
                            logging.warning(f"Button '{button_name}' not found, trying keyboard shortcut...")
                            try:
                                VK_Y = 0x59
                                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, VK_Y, 0)
                                time.sleep(0.05)
                                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, VK_Y, 0)
                                logging.info(f"✓ Sent 'Y' key to dialog")
                                button_clicked = True
                            except Exception as e:
                                logging.warning(f"Could not send key: {e}")
                        if button_clicked:
                            time.sleep(0.3)
                            return True
                    except Exception as e:
                        logging.warning(f"Error handling dialog: {e}")
            time.sleep(0.2)
        if not dialog_handled:
            logging.warning(f"Dialog not found within {timeout} seconds (may have been handled or not appeared)")
        return False
    
    # Start monitoring in background thread
    thread = threading.Thread(target=monitor_and_click, daemon=True)
    thread.start()
    return thread

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
    Automate Excel: load JSON, write to sheet, run macro, collect results, and save output.
    Handles error logging and cleanup.
    """
    try:
        with open(JSON_INPUT_PATH, "r", encoding="utf-8") as f:
            all_data = json.load(f)
        input_data = all_data[:MAX_RECORDS]
        logging.info(f"Loaded {len(input_data)} records (out of {len(all_data)} total)")
    except Exception as e:
        logging.error(f"Error loading input JSON: {e}")
        return

    unblock_file(os.path.abspath(EXCEL_PATH))

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = True
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = True
    excel.EnableEvents = True
    try:
        try:
            excel.AutomationSecurity = 1
        except Exception:
            logging.warning("Could not set AutomationSecurity (may require admin rights)")

        wb = None
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
            in_sheet = wb.Sheets(INPUT_SHEET)
        except Exception:
            alt_sheet = "Title" if INPUT_SHEET == "タイトル" else "タイトル"
            try:
                in_sheet = wb.Sheets(alt_sheet)
                logging.info(f"Using sheet '{alt_sheet}' instead of '{INPUT_SHEET}'")
            except Exception:
                logging.error(f"Could not find sheet '{INPUT_SHEET}' or '{alt_sheet}'")
                raise RuntimeError(f"Could not find sheet '{INPUT_SHEET}' or '{alt_sheet}'")
        try:
            out_sheet = wb.Sheets(OUTPUT_SHEET)
        except Exception:
            out_sheet = in_sheet

        for row in range(START_ROW, START_ROW + len(input_data) + 50):
            for col in INPUT_MAPPING.values():
                in_sheet.Range(f"{col}{row}").Value = ""
            in_sheet.Range(f"E{row}").Value = ""
            in_sheet.Range(f"G{row}").Value = ""
            in_sheet.Range(f"H{row}").Value = ""

        results = []
        logging.info(f"Writing {len(input_data)} records to columns A-D starting at row {START_ROW}...")
        for idx, row_data in enumerate(input_data):
            row = START_ROW + idx
            for key, col in INPUT_MAPPING.items():
                if key in row_data:
                    value = row_data[key]
                    in_sheet.Range(f"{col}{row}").Value = value
                    if idx < 3:
                        logging.info(f"Row {row}, Col {col}: {key} = {str(value)[:50]}...")
        logging.info(f"✓ Data written successfully to {len(input_data)} rows")

        logging.info(f"Writing formulas to columns G and H...")
        for idx in range(len(input_data)):
            row = START_ROW + idx
            next_row = row + 1
            value = idx + 1
            in_sheet.Range(f"E{row}").Value = value
            formula_g = f"=getvol(D{next_row})"
            in_sheet.Range(f"G{row}").Formula = formula_g
            formula_h = f"=GetPureTitle(D{row})"
            in_sheet.Range(f"H{row}").Formula = formula_h
            if idx < 3:
                logging.info(f"Row {row}: G={formula_g}, H={formula_h}")
        logging.info(f"✓ Formulas written successfully to {len(input_data)} rows")

        available_macros = []
        try:
            logging.info("Searching for Available Macros...")
            for vb_component in wb.VBProject.VBComponents:
                code_module = vb_component.CodeModule
                if code_module.CountOfLines > 0:
                    for i in range(1, code_module.CountOfLines + 1):
                        line = code_module.Lines(i, 1)
                        if line.strip().startswith(("Sub ", "Function ")):
                            proc_name = line.split()[1].split("(")[0]
                            available_macros.append((proc_name, vb_component.Name))
                            logging.info(f"Found: {proc_name} in module '{vb_component.Name}'")
            if not available_macros:
                logging.info("No macros found in VBA project")
            else:
                logging.info(f"Total macros found: {len(available_macros)}")
        except Exception as e:
            logging.warning(f"Could not list macros: {e}")
            logging.warning("IMPORTANT: You need to enable 'Trust access to the VBA project object model'")

        macro_called = False
        workbook_name = wb.Name
        workbook_base = os.path.basename(EXCEL_PATH)
        macro_attempts = [
            MACRO_NAME,
            f"{workbook_name}!{MACRO_NAME}",
            f"'{workbook_name}'!{MACRO_NAME}",
            f"'{workbook_base}'!{MACRO_NAME}",
        ]
        if available_macros:
            for macro_name, module_name in available_macros:
                if MACRO_NAME.lower() in macro_name.lower() or macro_name.lower() in MACRO_NAME.lower():
                    macro_attempts.insert(0, f"{module_name}.{macro_name}")
                    macro_attempts.insert(0, f"'{workbook_name}'!{module_name}.{macro_name}")

        logging.info(f"Attempting to Run Macro: {MACRO_NAME}")
        logging.info("Starting auto-click handler for VBA dialog...")
        dialog_thread = close_msgbox(window_title_contains="マクロ実行確認", button_name="はい", timeout=15)
        time.sleep(0.5)
        for macro_path in macro_attempts:
            try:
                logging.info(f"Trying: {macro_path}")
                excel.Run(macro_path)
                macro_called = True
                logging.info(f"✓ SUCCESS! Macro executed using: {macro_path}")
                break
            except Exception as e:
                error_msg = str(e)
                logging.warning(f"✗ Failed: {error_msg[:100]}")
                try:
                    wb.Application.Run(macro_path)
                    macro_called = True
                    logging.info(f"✓ SUCCESS! Macro executed using wb.Application.Run: {macro_path}")
                    break
                except Exception:
                    continue
        if not macro_called:
            try:
                logging.info("Trying alternative method: Application.OnTime...")
                wb.Activate()
                excel.Application.OnTime(
                    excel.Application.Now + excel.Application.TimeValue("00:00:01"),
                    f"'{workbook_name}'!{MACRO_NAME}"
                )
                time.sleep(2)
                macro_called = True
                logging.info("✓ Macro scheduled via OnTime")
            except Exception as e:
                logging.warning(f"✗ OnTime method failed: {e}")
        if not macro_called:
            logging.error(f"ERROR: Could not run macro '{MACRO_NAME}'")
            logging.error("TROUBLESHOOTING STEPS:")
            logging.error("1. CHECK EXCEL SECURITY SETTINGS:")
            logging.error("   - Open Excel manually")
            logging.error("   - File > Options > Trust Center > Trust Center Settings")
            logging.error("   - Macro Settings: Select 'Enable all macros' (or 'Enable macros with notification')")
            logging.error("   - Developer Macro Settings: CHECK 'Trust access to the VBA project object model'")
            logging.error("   - Click OK and restart Excel")
            logging.error("2. VERIFY MACRO EXISTS:")
            logging.error(f"   - Open {workbook_name} in Excel")
            logging.error("   - Press Alt+F8 to view macros")
            logging.error(f"   - Look for '{MACRO_NAME}'")
            if available_macros:
                logging.error(f"   Found macros in workbook:")
                for macro_name, module_name in available_macros:
                    logging.error(f"     - {macro_name} (in {module_name})")
            logging.error("3. UNBLOCK THE FILE:")
            logging.error(f"   - Right-click: {EXCEL_PATH}")
            logging.error("   - Properties > General tab")
            logging.error("   - If 'Unblock' checkbox exists, check it")
            logging.error("4. TRY RUNNING MACRO MANUALLY:")
            logging.error("   - Open the Excel file")
            logging.error("   - Press Alt+F8")
            logging.error(f"   - Run '{MACRO_NAME}' manually")
            logging.error("   - If it works manually but not from Python, it's a security issue")
            raise RuntimeError(f"Could not run macro '{MACRO_NAME}'. Please check Excel security settings and verify the macro exists.")
        logging.info("Waiting for macro to complete and dialog to be handled...")
        time.sleep(3)
        if dialog_thread.is_alive():
            logging.info("Dialog handler still running, waiting up to 5 more seconds...")
            dialog_thread.join(timeout=5)
        for idx, row_data in enumerate(input_data):
            row = START_ROW + idx
            record = {}
            for col, json_key in OUTPUT_MAPPING.items():
                cell = out_sheet.Range(f"{col}{row}")
                record[json_key] = cell.Value
                if(json_key == "タイトル"):
                    colors = cell.DisplayFormat.Interior.Color
                    if colors == 9895780.0:
                        record["color"] = True
                    else:
                        record["color"] = False
                if(json_key == "巻数"):
                    cell_value = cell.Value
                    if cell_value is None:
                        record["巻数"] = 0
                    else:
                        try:
                            record["巻数"] = int(cell_value)
                        except ValueError:
                            record["巻数"] = 0
            results.append(record)
        with open(JSON_OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)
        logging.info(f"Automation complete! → {JSON_OUTPUT_PATH}")
        return results
    except Exception as e:
        logging.error(f"Error during Excel automation: {e}")
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
