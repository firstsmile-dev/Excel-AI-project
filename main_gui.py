"""
main_gui.py - GUI for Excel-AI-project
Features: history window, timer, start/stop buttons, persistent window after completion.
"""
from __future__ import annotations
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import time
import os
import importlib
import logging
from typing import Sequence, Any
import importlib
import json
import win32com.client as win32
import glob
from dotenv import load_dotenv
from openai import APIError, AuthenticationError, OpenAI

# Load environment variables from .env file
load_dotenv()

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
# =============================
OUTPUT_MAPPING = {
    "I": "タイトル",
    "F": "ASIN",
    "G": "巻数",
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


def edit_json_with_openai(
    json_path: str,
    model: str = "gpt-4.1-mini",
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
    system_msg = """# Identity
                    あなたは、入力されたテキストからマンガ／ラノベ／書籍タイトルの「正式名称のみ」を抽出し、不要要素を取り除いて整形するアシスタントです。  
                    あなたの目的は、タイトル一覧をクリーンで一貫した形式に統一して出力することです。

                    # Instructions
                    以下のルールに厳密に従って出力してください。

                    1. 入力に含まれる **話数、巻数、出版社名、サブタイトル、記号、エロ・成人タグ、説明文、広告文、シリーズ名以外の情報** をすべて削除してください。  
                    2. 抽出対象は **作品タイトルの正式名称のみ** とします。  
                    3. **同一作品の表記ゆれ**（全角／半角、記号、サブタイトルの有無、略称の違い）は **一つの正式な表記に統一** してください。  
                    4. 出力は **1行につき1タイトル** とします。  
                    5. **タイトル以外の情報を推測して追加してはいけません。**  
                    6. 原作名とシリーズ名の区別が必要な場合は、**シリーズ名を優先** してください。  
                    7. 表記は **日本語のまま、正式名称に統一** してください。  
                    8. **コメントや説明文は一切書かず、タイトルのみを出力** してください。
                    9. タイトルに巻数を示す数字（3、ローマ数字、日本語の漢数字など）が含まれている場合はお知らせください。
                       数字が含まれているときは、それが本の巻数を正確に示しているかどうかを判定し、巻数であると判断した場合は数字だけを教えてください（例：3）。

                    # Example1
                    <user_query>
                    ちびっ子転生日記帳～お友達いっは?いつくりましゅ!～ THE COMIC 2 (マッグガーデンコミック Beat'sシリーズ)  
                    </user_query>

                    <assistant_response>
                    ちびっ子転生日記帳～お友達いっぱいつくりましゅ!～ THE COMIC
                    2
                    </assistant_response>

                    # Example2

                    <user_query>
                    ミッドナイトレストラン 7to7  
                    </user_query>

                    <assistant_response>
                    ミッドナイトレストラン 7to7
                    0
                    </assistant_response>

                    # Example3

                    <user_query>
                    ながたんと青と-いちかの料理帖-
                    </user_query>

                    <assistant_response>
                    ながたんと青と－いちかの料理帖－  
                    0
                    </assistant_response>

                    # Example4

                    <user_query>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～(コミック) :  
                    </user_query>

                    <assistant_response>
                    おっさん底辺治癒士と愛娘の辺境ライフ～中年男が回復スキルに覚醒して、英雄へ成り上がる～ 
                    0
                    </assistant_response>

                    # Example5

                    <user_query>
                    ハボウの轍 4 ~公安調査庁調査官・土師空也~
                    </user_query>

                    <assistant_response>
                    ハボウの轍～公安調査庁調査官・土師空也～
                    4
                    </assistant_response>

                    # Example6

                    <user_query>
                    バリタチNo.1に負けた俺がネコデビューするまで (DAITO COMICS)
                    </user_query>

                    <assistant_response>
                    バリタチNo.1に負けた俺がネコデビューするまで
                    0
                    </assistant_response>

                    # Example7

                    <user_query>
                    私と結婚した事、後悔していませんか?VI (秋水デジタルコミックス)
                    </user_query>

                    <assistant_response>
                    私と結婚した事、後悔していませんか?
                    4
                    </assistant_response>
                    
                    # Context
                    以下にユーザーが未整理のタイトル一覧を入力します。  
                    ルールに従って正式タイトルのみを抽出・整形してください。"""
    edited_data = []
    for item in data:
        user_content = item.get("タイトル")
        color = item.get("color", False)
        new_item = item.copy()
        try:
            if color == True:
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
                print("response text", text)
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
        with open(csv_path, "r", encoding="cp932", errors="ignore") as f:
            reader = csv.reader(f)
            rows = list(reader)
        real_data = rows[0:1]  # header row
        for item in json_data:
            row = [""] * len(rows[0])
            row[2] = item.get("タイトル", "") #C column
            row[6] = item.get("巻数", "") #G column
            row[14] = item.get("ASIN", "") #O column

            real_data.append(row)

        with open(csv_path, "w", encoding="cp932", newline="") as f:
            writer = csv.writer(f)
            writer.writerows(real_data)
        return True
    except Exception as exc:
        logging.error(f"Error converting JSON to CSV: {exc}")
        raise RuntimeError(f"Error converting JSON to CSV: {exc}") from exc


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Excel-AI Project")
        self.geometry("500x300")
        self.resizable(False, False)
        # History window
        history_frame = tk.LabelFrame(self, text="History", padx=5, pady=5)
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
        self.start_btn = ttk.Button(btn_frame, text="Start", command=self.start_workflow)
        self.start_btn.grid(row=0, column=0, padx=10)
        self.stop_btn = ttk.Button(btn_frame, text="Stop", command=self.stop_workflow, state="disabled")
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
        self.log_history("[START] Workflow started.")
        self._workflow_thread = threading.Thread(target=self.run_main_workflow, daemon=True)
        self._workflow_thread.start()

    def stop_workflow(self):
        self._timer_running = False
        self.start_btn.config(state="normal")
        self.stop_btn.config(state="disabled")
        self.log_history("[STOP] Process stopped by user.")
        messagebox.showinfo("Stopped", "Process stopped by user.")

    def update_timer(self):
        if self._timer_running:
            elapsed = int(time.time() - self._start_time)
            h, m = divmod(elapsed, 3600)
            m, s = divmod(m, 60)
            self.timer_var.set(f"{h:02}:{m:02}:{s:02}")
            self.after(500, self.update_timer)

    def run_main_workflow(self):
        try:
            self.log_history("[INFO] Running main workflow...")
            # Step 1: Run vba_simulation.py workflow
            vba_success = run_excel_process()
            if vba_success is not None:
                json_path = find_file_by_ext("./public", "json")
                csv_path = find_file_by_ext("./public", "csv")
                edited_data = edit_json_with_openai(json_path)
                convert_info = input_json_convert_csv(edited_data, csv_path)
            else:
                self.log_history("[ERROR] VBA simulation failed.")
                raise RuntimeError("VBA simulation failed.")
            self._timer_running = False
            self.stop_btn.config(state="disabled")
            self.start_btn.config(state="normal")
            self.log_history("[COMPLETE] Project workflow completed.")
            messagebox.showinfo("Completed", "Project workflow completed.")
        except Exception as e:
            self._timer_running = False
            self.stop_btn.config(state="disabled")
            self.start_btn.config(state="normal")
            self.log_history(f"[ERROR] {e}")
            messagebox.showerror("Error", f"Error: {e}")

    def on_close(self):
        if self._timer_running:
            if not messagebox.askokcancel("Quit", "Workflow is running. Quit anyway?"):
                return
        self.destroy()

if __name__ == "__main__":
    app = MainApp()
    app.mainloop()