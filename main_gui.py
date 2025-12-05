"""
main_gui.py - GUI for Excel-AI-project
Features: history window, timer, start/stop buttons, persistent window after completion.
"""
import tkinter as tk
from tkinter import ttk
from tkinter import messagebox
import threading
import time
import os
import importlib

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
            main_module = importlib.import_module("main")
            self.log_history("[INFO] Running main workflow...")
            main_module.main()
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