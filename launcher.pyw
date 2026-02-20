#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
TyranoScript Localization Kit -- GUI Launcher
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import subprocess
import threading
import os
import sys
import json
import re

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "config.json")
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PAUSE_FILE = os.path.join(SCRIPT_DIR, ".pause_signal")

ALL_PACKAGES = {
    "openpyxl": "openpyxl",
    "deep_translator": "deep-translator",
}

DEPS_EXTRACT = ["openpyxl"]
DEPS_TRANSLATE = ["openpyxl", "deep-translator"]
DEPS_APPLY = ["openpyxl"]


def _find_missing(pip_names):
    """Return list of pip package names that are not installed."""
    pip_to_import = {v: k for k, v in ALL_PACKAGES.items()}
    missing = []
    for pip_name in pip_names:
        import_name = pip_to_import.get(pip_name, pip_name)
        try:
            __import__(import_name)
        except ImportError:
            missing.append(pip_name)
    return missing


DEFAULT_SETTINGS = {
    "ks_dir": "",
    "xlsx": "translations.xlsx",
    "xlsx_translated": "translations_translated.xlsx",
    "out_dir": "translated",
    "engine": "google",
    "from_lang": "en",
    "to_lang": "ru",
    "deepl_key": "",
}


# -- Settings persistence ------------------------------------------------

def load_settings():
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            return {**DEFAULT_SETTINGS, **saved}
        except Exception:
            pass
    bat_path = os.path.join(SCRIPT_DIR, "config.bat")
    if os.path.exists(bat_path):
        try:
            with open(bat_path, "r", encoding="utf-8-sig") as f:
                text = f.read()
            m = re.search(r'set\s+"KS_DIR=(.+?)"', text)
            if m:
                settings = dict(DEFAULT_SETTINGS)
                settings["ks_dir"] = m.group(1)
                return settings
        except Exception:
            pass
    return dict(DEFAULT_SETTINGS)


def save_settings(settings):
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(settings, f, ensure_ascii=False, indent=2)
    bat_path = os.path.join(SCRIPT_DIR, "config.bat")
    bat_content = (
        '@echo off\r\n'
        'REM ============================================================\r\n'
        'REM   PROJECT SETTINGS (auto-updated by launcher.pyw)\r\n'
        'REM ============================================================\r\n'
        '\r\n'
        f'set "KS_DIR={settings["ks_dir"]}"\r\n'
        f'set "XLSX={settings["xlsx"]}"\r\n'
        f'set "XLSX_TRANSLATED={settings["xlsx_translated"]}"\r\n'
        f'set "OUT_DIR={settings["out_dir"]}"\r\n'
    )
    with open(bat_path, "wb") as f:
        f.write(bat_content.encode("utf-8"))


# -- Main window ----------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("TyranoScript Localization Kit")
        self.resizable(True, True)
        self.minsize(750, 560)
        self.geometry("800x640")

        self.settings = load_settings()
        self.running = False
        self.paused = False
        self._current_proc = None

        self._build_ui()
        self._center_window()
        self._cleanup_pause()

    def _center_window(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        x = (self.winfo_screenwidth() // 2) - (w // 2)
        y = (self.winfo_screenheight() // 2) - (h // 2)
        self.geometry(f"+{x}+{y}")

    def _cleanup_pause(self):
        """Remove stale pause file on startup."""
        if os.path.exists(PAUSE_FILE):
            os.remove(PAUSE_FILE)

    # -- UI ---------------------------------------------------------------

    def _build_ui(self):
        # Project path
        path_frame = ttk.LabelFrame(self, text="  Scenario Folder (.ks files)  ", padding=10)
        path_frame.pack(fill="x", padx=10, pady=(10, 5))

        self.path_var = tk.StringVar(value=self.settings["ks_dir"])
        ttk.Entry(path_frame, textvariable=self.path_var,
                  font=("Consolas", 10)).pack(side="left", fill="x", expand=True, padx=(0, 5))
        ttk.Button(path_frame, text="Browse...",
                   command=self._browse_folder).pack(side="right")

        # Translation settings
        trans_frame = ttk.LabelFrame(self, text="  Translation Settings  ", padding=10)
        trans_frame.pack(fill="x", padx=10, pady=5)

        # Row 1: engine + languages
        row1 = ttk.Frame(trans_frame)
        row1.pack(fill="x", pady=2)

        ttk.Label(row1, text="Engine:").pack(side="left", padx=(0, 5))
        self.engine_var = tk.StringVar(value=self.settings["engine"])
        ttk.Combobox(row1, textvariable=self.engine_var,
                     values=["google", "deepl", "libre", "mymemory"],
                     width=12, state="readonly").pack(side="left", padx=(0, 20))

        ttk.Label(row1, text="From:").pack(side="left", padx=(0, 5))
        self.from_var = tk.StringVar(value=self.settings["from_lang"])
        ttk.Entry(row1, textvariable=self.from_var, width=6).pack(side="left", padx=(0, 15))

        ttk.Label(row1, text="To:").pack(side="left", padx=(0, 5))
        self.to_var = tk.StringVar(value=self.settings["to_lang"])
        ttk.Entry(row1, textvariable=self.to_var, width=6).pack(side="left")

        ttk.Label(row1, text="   (en, ru, de, fr, ja, ko, zh, es...)",
                  foreground="gray").pack(side="left", padx=10)

        # Row 2: DeepL key
        row_deepl = ttk.Frame(trans_frame)
        row_deepl.pack(fill="x", pady=2)
        ttk.Label(row_deepl, text="DeepL API key:").pack(side="left", padx=(0, 5))
        self.deepl_key_var = tk.StringVar(value=self.settings.get("deepl_key", ""))
        ttk.Entry(row_deepl, textvariable=self.deepl_key_var,
                  width=45, show="*").pack(side="left", padx=(0, 10))
        ttk.Label(row_deepl, text="(free: deepl.com/pro-api)",
                  foreground="gray").pack(side="left")

        # Row 3: file names
        row2 = ttk.Frame(trans_frame)
        row2.pack(fill="x", pady=2)
        ttk.Label(row2, text="XLSX file:").pack(side="left", padx=(0, 5))
        self.xlsx_var = tk.StringVar(value=self.settings["xlsx"])
        ttk.Entry(row2, textvariable=self.xlsx_var, width=30).pack(side="left", padx=(0, 15))

        ttk.Label(row2, text="Output folder:").pack(side="left", padx=(0, 5))
        self.out_var = tk.StringVar(value=self.settings["out_dir"])
        ttk.Entry(row2, textvariable=self.out_var, width=20).pack(side="left")

        # Action buttons
        btn_frame = ttk.Frame(self, padding=5)
        btn_frame.pack(fill="x", padx=10, pady=5)

        self.btn_extract = ttk.Button(btn_frame, text="1. Extract",
                                      command=self._run_extract)
        self.btn_extract.pack(side="left", padx=5, expand=True, fill="x")

        self.btn_translate = ttk.Button(btn_frame, text="2. Translate",
                                        command=self._run_translate)
        self.btn_translate.pack(side="left", padx=5, expand=True, fill="x")

        self.btn_pause = ttk.Button(btn_frame, text="Pause",
                                    command=self._toggle_pause, state="disabled")
        self.btn_pause.pack(side="left", padx=5, fill="x")

        self.btn_apply = ttk.Button(btn_frame, text="3. Apply",
                                    command=self._run_apply)
        self.btn_apply.pack(side="left", padx=5, expand=True, fill="x")

        # Console
        console_frame = ttk.LabelFrame(self, text="  Log  ", padding=5)
        console_frame.pack(fill="both", expand=True, padx=10, pady=(5, 10))

        self.console = tk.Text(console_frame, wrap="word", font=("Consolas", 9),
                               bg="#1e1e1e", fg="#cccccc", insertbackground="#cccccc",
                               state="disabled", relief="flat")
        scrollbar = ttk.Scrollbar(console_frame, command=self.console.yview)
        self.console.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.console.pack(fill="both", expand=True)

        self.console.tag_configure("info", foreground="#569cd6")
        self.console.tag_configure("success", foreground="#6a9955")
        self.console.tag_configure("error", foreground="#f44747")
        self.console.tag_configure("header", foreground="#dcdcaa",
                                   font=("Consolas", 9, "bold"))

        # Status bar
        status_frame = ttk.Frame(self)
        status_frame.pack(fill="x", padx=10, pady=(0, 5))
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var,
                  foreground="gray").pack(side="left")
        ttk.Button(status_frame, text="Clear log",
                   command=self._clear_log).pack(side="right")

    # -- Helpers ----------------------------------------------------------

    def _browse_folder(self):
        current = self.path_var.get()
        initial = current if os.path.isdir(current) else ""
        folder = filedialog.askdirectory(
            title="Select scenario folder (.ks files)", initialdir=initial)
        if folder:
            folder = folder.replace("/", "\\")
            self.path_var.set(folder)
            self._log(f"Path set: {folder}\n", "info")

    def _gather_settings(self):
        return {
            "ks_dir": self.path_var.get().strip(),
            "xlsx": self.xlsx_var.get().strip() or "translations.xlsx",
            "xlsx_translated": self.settings.get("xlsx_translated",
                                                  "translations_translated.xlsx"),
            "out_dir": self.out_var.get().strip() or "translated",
            "engine": self.engine_var.get(),
            "from_lang": self.from_var.get().strip() or "en",
            "to_lang": self.to_var.get().strip() or "ru",
            "deepl_key": self.deepl_key_var.get().strip(),
        }

    def _save_current(self):
        self.settings = self._gather_settings()
        save_settings(self.settings)

    def _log(self, text, tag=None):
        self.console.configure(state="normal")
        if tag:
            self.console.insert("end", text, tag)
        else:
            self.console.insert("end", text)
        self.console.see("end")
        self.console.configure(state="disabled")

    def _clear_log(self):
        self.console.configure(state="normal")
        self.console.delete("1.0", "end")
        self.console.configure(state="disabled")

    def _set_buttons(self, enabled):
        state = "normal" if enabled else "disabled"
        self.btn_extract.configure(state=state)
        self.btn_translate.configure(state=state)
        self.btn_apply.configure(state=state)
        self.btn_pause.configure(state="disabled")
        self.running = not enabled
        if enabled:
            self.paused = False
            self._current_proc = None
            self._cleanup_pause()

    # -- Pause / Resume ---------------------------------------------------

    def _toggle_pause(self):
        if not self.paused:
            # Pause: create signal file
            try:
                with open(PAUSE_FILE, "w") as f:
                    f.write("paused")
            except Exception:
                pass
            self.paused = True
            self.btn_pause.configure(text="Resume")
            self.status_var.set("Paused")
            self._log("--- PAUSED (waiting for current batch to finish) ---\n", "info")
        else:
            # Resume: remove signal file
            self._cleanup_pause()
            self.paused = False
            self.btn_pause.configure(text="Pause")
            self.status_var.set("Resumed...")
            self._log("--- RESUMED ---\n", "info")

    # -- Run command ------------------------------------------------------

    def _run_command(self, args, label, deps=None, allow_pause=False):
        """Run a subprocess in a thread, streaming output to the console."""
        self._save_current()
        self._set_buttons(False)
        if allow_pause:
            self.btn_pause.configure(state="normal", text="Pause")
            self.paused = False
        self.status_var.set(f"Running: {label}...")

        self._log(f"\n{'='*60}\n", "header")
        self._log(f"  {label}\n", "header")
        self._log(f"{'='*60}\n\n", "header")

        def worker():
            try:
                # Auto-install missing dependencies
                if deps:
                    missing = _find_missing(deps)
                    if missing:
                        self.after(0, self._log,
                                   f"Installing: {', '.join(missing)}...\n", "info")
                        pip_cmd = [sys.executable, "-m", "pip", "install"] + missing
                        pip_proc = subprocess.Popen(
                            pip_cmd,
                            stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                            text=True, encoding="utf-8", errors="replace",
                            cwd=SCRIPT_DIR,
                            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                        )
                        for line in pip_proc.stdout:
                            self.after(0, self._log, line)
                        pip_proc.wait()
                        if pip_proc.returncode != 0:
                            self.after(0, self._log,
                                       "\nFailed to install packages!\n"
                                       f"Try manually: pip install {' '.join(missing)}\n",
                                       "error")
                            self.after(0, self.status_var.set,
                                       "Error: pip install failed")
                            return
                        self.after(0, self._log, "Packages OK\n\n", "success")

                # Run the actual command
                run_args = list(args)
                if run_args and run_args[0] == sys.executable:
                    run_args.insert(1, "-u")
                proc = subprocess.Popen(
                    run_args,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, encoding="utf-8", errors="replace",
                    cwd=SCRIPT_DIR,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
                self._current_proc = proc

                for line in proc.stdout:
                    self.after(0, self._log, line)
                proc.wait()

                if proc.returncode == 0:
                    self.after(0, self._log, "\nDone!\n", "success")
                    self.after(0, self.status_var.set, f"{label} -- done!")
                else:
                    self.after(0, self._log,
                               f"\nError (code {proc.returncode})\n", "error")
                    self.after(0, self.status_var.set, f"{label} -- error!")
            except FileNotFoundError:
                self.after(0, self._log,
                           "\nPython not found! Make sure Python is installed "
                           "and added to PATH.\n", "error")
                self.after(0, self.status_var.set, "Error: Python not found")
            except Exception as e:
                self.after(0, self._log, f"\n{e}\n", "error")
                self.after(0, self.status_var.set, "Error!")
            finally:
                self.after(0, self._set_buttons, True)

        threading.Thread(target=worker, daemon=True).start()

    # -- Actions ----------------------------------------------------------

    def _run_extract(self):
        s = self._gather_settings()
        if not s["ks_dir"]:
            messagebox.showwarning("Error",
                                   "Set the path to scenario folder first!")
            return
        self._run_command(
            [sys.executable, os.path.join(SCRIPT_DIR, "tyrano_l10n.py"),
             "extract", s["ks_dir"], "--output", s["xlsx"]],
            "Extract strings",
            deps=DEPS_EXTRACT,
        )

    def _run_translate(self):
        s = self._gather_settings()
        xlsx = s["xlsx"]
        if not os.path.exists(os.path.join(SCRIPT_DIR, xlsx)):
            if not os.path.exists(xlsx):
                messagebox.showwarning("Error",
                    f"File {xlsx} not found.\nRun step 1 (Extract) first.")
                return

        cmd = [sys.executable, os.path.join(SCRIPT_DIR, "translate_xlsx.py"),
               xlsx, "--engine", s["engine"],
               "--from", s["from_lang"], "--to", s["to_lang"]]

        if s["engine"] == "deepl" and s["deepl_key"]:
            cmd += ["--deepl-key", s["deepl_key"]]

        self._run_command(cmd,
            f"Translate ({s['engine']}: {s['from_lang']}->{s['to_lang']})",
            deps=DEPS_TRANSLATE,
            allow_pause=True,
        )

    def _run_apply(self):
        s = self._gather_settings()
        if not s["ks_dir"]:
            messagebox.showwarning("Error",
                                   "Set the path to scenario folder first!")
            return
        self._run_command(
            [sys.executable, os.path.join(SCRIPT_DIR, "tyrano_l10n.py"),
             "apply", s["ks_dir"], s["xlsx_translated"],
             "--output", s["out_dir"]],
            "Apply translations",
            deps=DEPS_APPLY,
        )


# -- Entry point ----------------------------------------------------------

if __name__ == "__main__":
    app = App()
    app.mainloop()
