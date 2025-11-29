import json
import win32com.client as win32
import time
import os
import threading
import win32gui
import win32con

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
    "I": "指定文字列を削除/変換",
    "J": "削除語",
    "K": "特殊文字",
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
            
            # Use EnumWindows to find windows by substring match
            def enum_callback(hwnd, windows):
                if win32gui.IsWindowVisible(hwnd):
                    try:
                        title = win32gui.GetWindowText(hwnd)
                        if window_title_contains in title or "マクロ" in title or "実行確認" in title:
                            windows.append((hwnd, title))
                    except:
                        pass
                return True
            
            win32gui.EnumWindows(enum_callback, windows)
            
            # Process found windows
            for hwnd, title in windows:
                if not dialog_handled:
                    print(f"  Found dialog: {title}")
                    dialog_handled = True
                    
                    try:
                        # Bring dialog to foreground
                        win32gui.SetForegroundWindow(hwnd)
                        win32gui.BringWindowToTop(hwnd)
                        time.sleep(0.2)
                        
                        # Find and click the button
                        button_clicked = False
                        
                        def find_button(child_hwnd, param):
                            nonlocal button_clicked
                            try:
                                child_text = win32gui.GetWindowText(child_hwnd)
                                # Check if this is the button we want
                                if button_name in child_text:
                                    # Try to click the button
                                    win32gui.SendMessage(child_hwnd, win32con.BM_CLICK, 0, 0)
                                    print(f"  ✓ Clicked button: {child_text}")
                                    button_clicked = True
                                    return False  # Stop enumeration
                            except:
                                pass
                            return True  # Continue enumeration
                        
                        # Enumerate child windows to find the button
                        win32gui.EnumChildWindows(hwnd, find_button, None)
                        
                        # If button not found by text, try sending 'Y' key as fallback
                        if not button_clicked:
                            print(f"  Button '{button_name}' not found, trying keyboard shortcut...")
                            try:
                                # Send 'Y' key to dialog (common for Yes button)
                                VK_Y = 0x59
                                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, VK_Y, 0)
                                time.sleep(0.05)
                                win32gui.PostMessage(hwnd, win32con.WM_KEYUP, VK_Y, 0)
                                print(f"  ✓ Sent 'Y' key to dialog")
                                button_clicked = True
                            except Exception as e:
                                print(f"  ⚠ Could not send key: {e}")
                        
                        if button_clicked:
                            time.sleep(0.3)  # Give dialog time to close
                            return True
                            
                    except Exception as e:
                        print(f"  ⚠ Error handling dialog: {e}")
            
            time.sleep(0.2)  # Check every 200ms
        
        if not dialog_handled:
            print(f"  ⚠ Dialog not found within {timeout} seconds (may have been handled or not appeared)")
        return False
    
    # Start monitoring in background thread
    thread = threading.Thread(target=monitor_and_click, daemon=True)
    thread.start()
    return thread

def unblock_file(file_path):
    """Remove 'Mark of the Web' from file if it exists"""
    try:
        import subprocess
        # Use PowerShell to unblock the file
        cmd = f'powershell -Command "Unblock-File -Path \'{file_path}\'"'
        subprocess.run(cmd, shell=True, capture_output=True)
    except:
        pass

def run_excel_process():
    # Load JSON list
    with open(JSON_INPUT_PATH, "r", encoding="utf-8") as f:
        all_data = json.load(f)
    
    # Limit to MAX_RECORDS for testing
    input_data = all_data[:MAX_RECORDS]
    print(f"Loaded {len(input_data)} records (out of {len(all_data)} total)")

    # Unblock the Excel file (removes security warning)
    unblock_file(os.path.abspath(EXCEL_PATH))

    excel = win32.Dispatch("Excel.Application")
    excel.Visible = True  # Make visible to see security prompts and debug
    excel.DisplayAlerts = True  # Allow alerts so user can enable macros if prompted
    excel.AskToUpdateLinks = False
    excel.ScreenUpdating = True
    excel.EnableEvents = True
    
    # Try to set automation security (may not work on all Excel versions)
    try:
        excel.AutomationSecurity = 1  # msoAutomationSecurityLow
    except:
        print("Note: Could not set AutomationSecurity (may require admin rights)")

    wb = None
    try:
        excel_file_path = os.path.abspath(EXCEL_PATH)
        print(f"Opening Excel file: {excel_file_path}")
        
        # Open workbook - user may need to click "Enable Macros" if prompted
        wb = excel.Workbooks.Open(
            excel_file_path,
            UpdateLinks=0,  # Don't update links
            ReadOnly=False,
            Format=None,
            CorruptLoad=0
        )
        print("Workbook opened successfully")
        
        # Try to get the sheet (handle both "Title" and "タイトル")
        try:
            in_sheet = wb.Sheets(INPUT_SHEET)
        except:
            # Try alternative sheet name
            alt_sheet = "Title" if INPUT_SHEET == "タイトル" else "タイトル"
            try:
                in_sheet = wb.Sheets(alt_sheet)
                print(f"Using sheet '{alt_sheet}' instead of '{INPUT_SHEET}'")
            except:
                raise RuntimeError(f"Could not find sheet '{INPUT_SHEET}' or '{alt_sheet}'")
        
        try:
            out_sheet = wb.Sheets(OUTPUT_SHEET)
        except:
            out_sheet = in_sheet  # Use same sheet if output sheet not found

        # Clear previous rows (optional)
        for row in range(START_ROW, START_ROW + len(input_data) + 50):
            for col in INPUT_MAPPING.values():
                in_sheet.Range(f"{col}{row}").Value = ""
            # Also clear columns G and H
            in_sheet.Range(f"E{row}").Value = ""
            in_sheet.Range(f"G{row}").Value = ""
            in_sheet.Range(f"H{row}").Value = ""

        results = []

        # Write each row
        print(f"\nWriting {len(input_data)} records to columns A-D starting at row {START_ROW}...")
        for idx, row_data in enumerate(input_data):
            row = START_ROW + idx

            for key, col in INPUT_MAPPING.items():
                if key in row_data:
                    value = row_data[key]
                    in_sheet.Range(f"{col}{row}").Value = value
                    if idx < 3:  # Show first 3 rows for debugging
                        print(f"  Row {row}, Col {col}: {key} = {str(value)[:50]}...")
        
        print(f"✓ Data written successfully to {len(input_data)} rows")
        
        # Write formulas to columns G and H
        print(f"\nWriting formulas to columns G and H...")
        for idx in range(len(input_data)):
            row = START_ROW + idx
            next_row = row + 1
            
            value = idx + 1
            in_sheet.Range(f"E{row}").Value = value
            # Column G: =getvol(D{next_row})
            formula_g = f"=getvol(D{next_row})"
            in_sheet.Range(f"G{row}").Formula = formula_g
            
            # Column H: =GetPureTitle(D{current_row})
            formula_h = f"=GetPureTitle(D{row})"
            in_sheet.Range(f"H{row}").Formula = formula_h
            
            if idx < 3:  # Show first 3 rows for debugging
                print(f"  Row {row}: G={formula_g}, H={formula_h}")
        
        print(f"✓ Formulas written successfully to {len(input_data)} rows")

        # List available macros for debugging
        available_macros = []
        try:
            print("\n=== Searching for Available Macros ===")
            for vb_component in wb.VBProject.VBComponents:
                code_module = vb_component.CodeModule
                if code_module.CountOfLines > 0:
                    # Try to find procedures
                    for i in range(1, code_module.CountOfLines + 1):
                        line = code_module.Lines(i, 1)
                        if line.strip().startswith(("Sub ", "Function ")):
                            proc_name = line.split()[1].split("(")[0]
                            available_macros.append((proc_name, vb_component.Name))
                            print(f"  Found: {proc_name} in module '{vb_component.Name}'")
            
            if not available_macros:
                print("  No macros found in VBA project")
            else:
                print(f"\n  Total macros found: {len(available_macros)}")
        except Exception as e:
            print(f"⚠ Could not list macros: {e}")
            print("\n⚠ IMPORTANT: You need to enable 'Trust access to the VBA project object model'")
            print("   Steps:")
            print("   1. Open Excel manually")
            print("   2. File > Options > Trust Center > Trust Center Settings")
            print("   3. Macro Settings > Enable all macros (or at least 'Enable macros with notification')")
            print("   4. Developer Macro Settings > ✓ Trust access to the VBA project object model")
            print("   5. Restart Excel and try again")

        # Run macro ONCE for performance
        # Try different ways to call the macro
        macro_called = False
        
        # Build macro name variations
        workbook_name = wb.Name
        workbook_base = os.path.basename(EXCEL_PATH)
        macro_attempts = [
            MACRO_NAME,  # Simple name
            f"{workbook_name}!{MACRO_NAME}",  # Workbook!MacroName
            f"'{workbook_name}'!{MACRO_NAME}",  # 'Workbook'!MacroName
            f"'{workbook_base}'!{MACRO_NAME}",  # 'filename.xlsm'!MacroName
        ]
        
        # If we found macros, try module-specific names
        if available_macros:
            for macro_name, module_name in available_macros:
                if MACRO_NAME.lower() in macro_name.lower() or macro_name.lower() in MACRO_NAME.lower():
                    macro_attempts.insert(0, f"{module_name}.{macro_name}")
                    macro_attempts.insert(0, f"'{workbook_name}'!{module_name}.{macro_name}")
        
        print(f"\n=== Attempting to Run Macro: {MACRO_NAME} ===")
        
        # Start auto-click thread to handle VBA dialog BEFORE running macro
        print("  Starting auto-click handler for VBA dialog...")
        dialog_thread = close_msgbox(window_title_contains="マクロ実行確認", button_name="はい", timeout=15)
        time.sleep(0.5)  # Brief pause to ensure thread is ready
        
        # Try different execution methods
        for macro_path in macro_attempts:
            try:
                print(f"  Trying: {macro_path}")
                excel.Run(macro_path)
                macro_called = True
                print(f"  ✓ SUCCESS! Macro executed using: {macro_path}")
                break
            except Exception as e:
                error_msg = str(e)
                print(f"  ✗ Failed: {error_msg[:100]}")
                
                # Try with wb.Application.Run
                try:
                    wb.Application.Run(macro_path)
                    macro_called = True
                    print(f"  ✓ SUCCESS! Macro executed using wb.Application.Run: {macro_path}")
                    break
                except Exception as e2:
                    continue
        
        # Last resort: Try using SendKeys or OnTime (less reliable)
        if not macro_called:
            try:
                print(f"\n  Trying alternative method: Application.OnTime...")
                wb.Activate()
                excel.Application.OnTime(
                    excel.Application.Now + excel.Application.TimeValue("00:00:01"),
                    f"'{workbook_name}'!{MACRO_NAME}"
                )
                time.sleep(2)
                macro_called = True
                print(f"  ✓ Macro scheduled via OnTime")
            except Exception as e:
                print(f"  ✗ OnTime method failed: {e}")
        
        if not macro_called:
            print(f"\n❌ ERROR: Could not run macro '{MACRO_NAME}'")
            print("\n" + "="*60)
            print("TROUBLESHOOTING STEPS:")
            print("="*60)
            print("\n1. CHECK EXCEL SECURITY SETTINGS:")
            print("   - Open Excel manually")
            print("   - File > Options > Trust Center > Trust Center Settings")
            print("   - Macro Settings:")
            print("     → Select 'Enable all macros' (or 'Enable macros with notification')")
            print("   - Developer Macro Settings:")
            print("     → ✓ CHECK 'Trust access to the VBA project object model'")
            print("   - Click OK and restart Excel")
            print("\n2. VERIFY MACRO EXISTS:")
            print(f"   - Open {workbook_name} in Excel")
            print("   - Press Alt+F8 to view macros")
            print(f"   - Look for '{MACRO_NAME}'")
            if available_macros:
                print(f"\n   Found macros in workbook:")
                for macro_name, module_name in available_macros:
                    print(f"     - {macro_name} (in {module_name})")
            print("\n3. UNBLOCK THE FILE:")
            print(f"   - Right-click: {EXCEL_PATH}")
            print("   - Properties > General tab")
            print("   - If 'Unblock' checkbox exists, check it")
            print("\n4. TRY RUNNING MACRO MANUALLY:")
            print("   - Open the Excel file")
            print("   - Press Alt+F8")
            print(f"   - Run '{MACRO_NAME}' manually")
            print("   - If it works manually but not from Python, it's a security issue")
            print("="*60)
            raise RuntimeError(f"Could not run macro '{MACRO_NAME}'. Please check Excel security settings and verify the macro exists.")
        
        # Wait for dialog to be handled and macro to complete
        print("\n  Waiting for macro to complete and dialog to be handled...")
        # Wait a bit longer to ensure dialog appears and is handled
        time.sleep(3)  # Give time for dialog to appear and be auto-clicked
        
        # Optionally wait for dialog thread to complete (with timeout)
        if dialog_thread.is_alive():
            print("  Dialog handler still running, waiting up to 5 more seconds...")
            dialog_thread.join(timeout=5)

        # Collect results
        # for idx, row_data in enumerate(input_data):
        #     row = START_ROW + idx
        #     record = {}

        #     for col, json_key in OUTPUT_MAPPING.items():
        #         record[json_key] = out_sheet.Range(f"{col}{row}").Value

        #     results.append(record)

        for idx, row_data in enumerate(input_data):
            row = START_ROW + idx
            record = {}

            for col, json_key in OUTPUT_MAPPING.items():
                cell = out_sheet.Range(f"{col}{row}")

                record[json_key] = cell.Value
                record[f"{json_key}_color"] = cell.Interior.Color  # 例: 16777215

            results.append(record)

        # Save JSON output
        with open(JSON_OUTPUT_PATH, "w", encoding="utf-8") as f:
            json.dump(results, f, ensure_ascii=False, indent=2)

        print("Automation complete! →", JSON_OUTPUT_PATH)
        return results
    
    except Exception as e:
        print(f"Error during Excel automation: {e}")
        import traceback
        traceback.print_exc()
        raise
    finally:
        # Clean up Excel objects
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except:
                pass
        try:
            excel.Quit()
        except:
            pass


if __name__ == "__main__":
    run_excel_process()
