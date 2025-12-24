"""
main_gui.py - GUI for Excel-AI-project
Features: history window, timer, start/stop buttons, persistent window after completion.
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
from datetime import datetime
import threading
import time
import os
import logging
from typing import Sequence, Any
import json
import win32com.client as win32
import glob
from dotenv import load_dotenv
from openai import APIError, AuthenticationError, OpenAI
import sys
import pythoncom

# Load environment variables from .env file
load_dotenv()

# === SETTINGS ===
# Get the directory where this script is located
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# Fix: Always use the real public folder next to the .exe for all file writes/reads
if getattr(sys, 'frozen', False):
    # Running as .exe
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PUBLIC_DIR = os.path.join(BASE_DIR, 'public')

# Utility to find first file by extension in a folder
# Fix: Use sys._MEIPASS for PyInstaller compatibility

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller bundle."""
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath(os.path.dirname(__file__)), relative_path)

def find_file_by_ext(folder: str, ext: str) -> str | None:
    files = glob.glob(os.path.join(folder, f"*.{ext}"))
    return files[0] if files else None

# Update all file paths to use PUBLIC_DIR
EXCEL_PATH = find_file_by_ext(PUBLIC_DIR, "xlsm")
JSON_OUTPUT_PATH = os.path.join(PUBLIC_DIR, "target_macro_output.json")
CSV_OUTPUT_PATH = find_file_by_ext(PUBLIC_DIR, "csv")

INPUT_SHEET = "タイトル"  # Sheet name (can be "Title" or "タイトル")
OUTPUT_SHEET = "タイトル"
MACRO_NAME = "Trimming"

START_ROW = 2  # Excel input starts at row 2
MAX_RECORDS = 8000  # Limit to 10 records for testing
# =============================
OUTPUT_MAPPING = {
    "D": "Amazonタイトル",
    "C": "b_タイトル",
    "I": "タイトル",
    "F": "ASIN",
    "G": "巻数",
    "E": "b_巻数",
}

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

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
    pythoncom.CoInitialize()  # Ensure COM is initialized in this thread
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
            valid_color = False 
            for col, json_key in OUTPUT_MAPPING.items():
                cell = out_sheet.Range(f"{col}{row}")
                value = cell.Value
                record[json_key] = value
                if value not in (None, ""):
                    empty_row = False
                # --- Volume number (巻数) ---
                if json_key == "巻数":
                    if value is None:
                        valid_color = True
                        record["巻数"] = 1
                    else:
                        try:
                            record["巻数"] = int(value)
                        except:
                            valid_color = True
                            record["巻数"] = 1
                # --- Title color check ---
                if json_key == "タイトル":
                    color_value = cell.DisplayFormat.Interior.Color
                    valid_color = (color_value == 9895780.0)
                    if value is None or value == "":
                        valid_color = True
                if json_key == "b_タイトル":
                    if record["b_タイトル"] is None or record["b_タイトル"] == "":
                        valid_color = True
                if json_key == "b_巻数":
                    if record["b_巻数"] != record["巻数"] or record["b_巻数"] == "" or record["b_巻数"] is None:
                        valid_color = True
            if empty_row:
                break
            if valid_color:
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


def edit_json_with_openai(
    json_path: str,
    model: str = "gpt-4-turbo",
    api_key: str | None = None,
) -> Any:
    """
    Send JSON data to OpenAI for processing and return the edited result.
    Handles API key retrieval, error handling, and logging.
    """
    # Get API key from parameter, .env file, environment variable, or raise error
    if api_key:
        api_key_value = api_key
    else:
        api_key_value = os.getenv("OPENAI_API_KEY")
        if not api_key_value:
            logging.error("OpenAI API key not provided. Set it as a parameter, or set OPENAI_API_KEY in your .env file or environment variable.")
            raise ValueError(
                "OpenAI API key not provided. Set it as a parameter, or set OPENAI_API_KEY in your .env file or environment variable."
            )

    # Initialize OpenAI client
    client = OpenAI(api_key=api_key_value)

    # Load data
    try:
        with open(json_path, "r", encoding="utf-8") as file_handle:
            data = json.load(file_handle)
        logging.info(f"Loaded JSON data from {json_path}")
    except FileNotFoundError as exc:
        logging.error(f"JSON file not found: {json_path}")
        raise FileNotFoundError(f"JSON file not found: {json_path}") from exc
    except json.JSONDecodeError as exc:
        logging.error(f"Invalid JSON in file {json_path}: {exc}")
        raise json.JSONDecodeError(
            f"Invalid JSON in file {json_path}: {exc}", exc.doc, exc.pos
        ) from exc

    # Compose system message and user content
    system_msg = os.getenv("SYSTEM_PROMPT")
    edited_data = []
    for item in data:
        user_content = item.get("タイトル")
        # Fix: Ensure user_content is always a string
        if user_content is None or user_content == "":
            user_content = item.get("Amazonタイトル")
        else:
            user_content = str(user_content)
        new_item = item.copy()
        try:
            response = client.responses.create(
                model=model,
                instructions=system_msg,
                input=[
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "input_text",
                                "text": user_content,
                            }
                        ],
                    }
                ],
            )
            text = response.output_text
            print(text)
            if(lines := text.split("\n")) and len(lines) >= 2:
                lines = [line.strip() for line in text.split("\n") if line.strip()]
                title = lines[0]
                explanation = lines[1]
                new_item["タイトル"] = title
                if(explanation != "0" ):
                    new_item["巻数"] = explanation
            else:
                new_item["タイトル"] = user_content
            edited_data.append(new_item)
        except json.JSONDecodeError as exc:
            logging.error(f"Model did not return valid JSON: {exc}")
            raise ValueError(
                f"Model did not return valid JSON: {exc}"
            ) from exc
        except AuthenticationError as exc:
            logging.error("Invalid OpenAI API key. Please check your API key.")
            raise ValueError("Invalid OpenAI API key. Please check your API key.") from exc
        except APIError as exc:
            logging.error(f"OpenAI API error: {exc}")
            raise RuntimeError(f"OpenAI API error: {exc}") from exc

    return edited_data

def input_json_convert_csv(json_data, csv_path:str):
    import csv
    """Convert JSON data to CSV and save to the specified path."""
    if not json_data:
        logging.warning("No data provided for CSV conversion.")
        return
    try:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"結果CSV_{timestamp}.csv"
        header_rows = list(os.getenv("HEADER_ROWS").split(",")) if os.getenv("HEADER_ROWS") else []
        print("rows", header_rows)
        real_data = [header_rows]
        print("real_data", real_data)
        for item in json_data:
            row = [""] * len(header_rows)
            if len(row) > 2:
                row[2] = str(item.get("タイトル", "")) #C column
            if len(row) > 6:
                row[6] = str(item.get("巻数", "")) #G column
            if len(row) > 14:
                row[14] = str(item.get("ASIN", "")) #O column
            real_data.append(row)
        print("real_data", real_data)
        # Fix: Use 'utf-8-sig' encoding for writing CSV to support all characters
        result_path = os.path.join(PUBLIC_DIR, filename)
        with open(result_path, "w", encoding="cp932", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(real_data)
            print("CSV saved to:", result_path)
        return True
    except Exception as exc:
        logging.error(f"Error converting JSON to CSV: {exc}")
        raise RuntimeError(f"Error converting JSON to CSV: {exc}") from exc


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-AIプロジェクト")
        self.geometry("500x300")
        self.resizable(False, False)
        # History window
        history_frame = tk.LabelFrame(self, text="歴史", padx=5, pady=5)
        history_frame.pack(fill="both", expand=False, padx=10, pady=10)
        self.history_text = tk.Text(history_frame, height=10, state="disabled", wrap="word")
        self.history_text.pack(fill="both", expand=True)
        # Timer
        self.timer_var = tk.StringVar(value="00:00:00")
        self.timer_label = tk.Label(self, textvariable=self.timer_var, font=("Arial", 18))
        self.timer_label.pack(pady=10)
        # Buttons
        btn_frame = tk.Frame(self)
        btn_frame.pack(pady=20)
        self.start_btn = ttk.Button(btn_frame, text="開始", command=self.start_workflow)
        self.start_btn.grid(row=0, column=0, padx=10)
        self.stop_btn = ttk.Button(btn_frame, text="停止", command=self.stop_workflow, state="disabled")
        self.stop_btn.grid(row=0, column=1, padx=10)
        # State
        self._timer_running = False
        self._workflow_thread = None
        self._start_time = None
        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def log_history(self, message):
        self.history_text.config(state="normal")
        self.history_text.insert("end", message + "\n")
        self.history_text.see("end")
        self.history_text.config(state="disabled")

    def start_workflow(self):
        self.start_btn.config(state="disabled")
        self.stop_btn.config(state="normal")
        self._timer_running = True
        self._start_time = time.time()
        self.update_timer()
        self.log_history("[START] ワークフローが開始されました。")
        self._workflow_thread = threading.Thread(target=self.run_main_workflow, daemon=True)
        self._workflow_thread.start()

    def stop_workflow(self):
        self._timer_running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.log_history("[STOP] ユーザーによってプロセスが停止されました。")
        messagebox.showinfo("停止", "プロセスはユーザーによって停止されました。")

    def update_timer(self):
        if self._timer_running:
            elapsed = int(time.time() - self._start_time)
            h, m = divmod(elapsed, 3600)
            m, s = divmod(m, 60)
            self.timer_var.set(f"{h:02}:{m:02}:{s:02}")
            self.after(500, self.update_timer)

    def run_main_workflow(self):
        try:
            self.log_history("[INFO] メインワークフローを実行しています...")
            # Step 1: Run vba_simulation.py workflow
            vba_success = run_excel_process()
            if vba_success is not None:
                edited_data = edit_json_with_openai(JSON_OUTPUT_PATH)
                convert_info = input_json_convert_csv(edited_data, CSV_OUTPUT_PATH)
            else:
                self.log_history("[エラー] VBA シミュレーションに失敗しました。")
                raise RuntimeError("VBA simulation failed.")
            if(convert_info):
                self.log_history("[INFO] CSV 変換が正常に完了しました。")
                self._timer_running = False
                self.stop_btn.config(state="disabled")
                self.start_btn.config(state="normal")
                self.log_history("[完了] プロジェクト ワークフローが完了しました。")
                messagebox.showinfo("完了", "プロジェクトワークフローが完了しました。")
        except Exception as e:
            self._timer_running = False
            self.stop_btn.config(state="disabled")
            self.start_btn.config(state="normal")
            self.log_history(f"[ERROR] {e}")
            messagebox.showerror("Error", f"Error: {e}")

    def on_close(self):
        if self._timer_running:
            if not messagebox.askokcancel("終了", "ワークフローは実行中です。終了しますか?"):
                return
        self.destroy()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()