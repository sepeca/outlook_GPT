#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Jeden panel: editor .env + spuštění hlavní aplikace (Outlook → GPT → Excel).
- Vyplň .env proměnné
- Klikni „Uložit & Spustit“ → otevře se nová konzole s interaktivním menu main.py

Spuštění: python gui_env.py
"""
import os
import sys
import re
from datetime import datetime
import threading
import subprocess
import tkinter as tk
from tkinter import ttk, messagebox, filedialog

try:
    from dotenv import dotenv_values
except Exception:
    dotenv_values = lambda p: {}

# --- persistent env path for frozen exe ---
# OLD:
# APP_DIR = os.path.dirname(os.path.abspath(__file__))
# ENV_PATH = os.path.join(APP_DIR, ".env")
APP_DIR = os.path.dirname(os.path.abspath(__file__))
_MEI_BASE = getattr(sys, "_MEIPASS", APP_DIR)  # resource base (for .env.example in exe)
if getattr(sys, "frozen", False):
    ENV_DIR = os.path.join(os.getenv("APPDATA"), "OutlookGPT")
    os.makedirs(ENV_DIR, exist_ok=True)
    ENV_PATH = os.path.join(ENV_DIR, ".env")
else:
    ENV_PATH = os.path.join(APP_DIR, ".env")
EXAMPLE_PATH = os.path.join(_MEI_BASE, ".env.example")  # bundled example if any

def _cli_cmd():
    """Return command to run the main pipeline."""
    if getattr(sys, "frozen", False):
        return [sys.executable, "--run-main"]
    return [sys.executable, os.path.join(APP_DIR, "main.py")]

# --- FIELDS extended with template, fetch sent, debug, etc. ---
FIELDS = [
    ("OPENAI_API_KEY", "OpenAI API klíč (sk-…)", "entry_secret", {}),
    ("OPENAI_BASE_URL", "OpenAI Base URL", "entry", {"default": "https://api.openai.com/v1"}),
    ("OPENAI_MODEL", "Model", "combo", {"values": ["gpt-4o-mini", "gpt-4o", "gpt-3.5-turbo"], "default": "gpt-4o-mini"}),

    ("MY_NAME", "Vase jmeno a prijmeni přes ','", "entry", {"default": "<NAME>"}),
    ("MY_EMAILS", "Vas/y email/y přes ','", "entry", {"default": "<EMAIL>"}),

    ("DATE_FROM", "Datum od (YYYY-MM-DD)", "entry", {"default": ""}),
    ("DATE_TO",   "Datum do (YYYY-MM-DD)", "entry", {"default": ""}),

    ("DAYS_BACK", "Výchozí počet dní zpět (fallback)", "entry", {"default": "7"}),
    ("MAX_EMAILS", "Max. počet e-mailů na běh", "entry", {"default": "200"}),
    ("OUTLOOK_FOLDER", "Složka Outlooku (např. Inbox/Subfolder)", "entry", {"default": "Inbox"}),
    ("STATUS", "Stav přečtení", "combo", {"values": ["all", "unread", "read"], "default": "all"}),
    # Output
    ("OUTPUT_DIR", "Složka pro uložení", "dir_browse", {"default": ""}),
    ("OUTPUT_NAME", "Název souboru (bez .xlsx)", "entry", {"default": "outlook_analysis"}),

    # template support
    ("TEMPLATE_XLSX", "Template XLSX", "entry_browse", {"default": ""}),
    ("TEMPLATE_SHEET", "Template sheet", "entry", {"default": ""}),
    ("TEMPLATE_START_AT_ROW3", "Start at row 3", "combo", {"values": ["true", "false"], "default": "true"}),

    # NEW: behavior flags
    ("FETCH_SENT_TOO", "Fetch Sent Items", "combo", {"values": ["true", "false"], "default": "true"}),
    ("DEBUG_GPT", "Debug GPT logs", "combo", {"values": ["true", "false"], "default": "false"}),


]


ACCEPTED_DATE_FORMATS = ["%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y"]

def _coerce_date_to_iso(s: str) -> str:
    """Return YYYY-MM-DD or empty."""
    s = (s or "").strip()
    if not s:
        return ""
    s = re.split(r"\s+", s, 1)[0] if " " in s else s
    for fmt in ACCEPTED_DATE_FORMATS:
        try:
            return datetime.strptime(s, fmt).strftime("%Y-%m-%d")
        except Exception:
            pass
    m = re.match(r"^(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4})", s)
    if m:
        d, mth, y = m.groups()
        return datetime(int(y), int(mth), int(d)).strftime("%Y-%m-%d")
    raise ValueError(f"Neplatné datum: {s}")

class EnvRunner(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook → GPT → Excel — Nastavení & Spuštění")
        self.geometry("900x650")
        self.minsize(840, 600)

        self.vars = {}
        self.proc = None

        self._build_ui()
        self._load_env()

    def _build_ui(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill="both", expand=True)

        tab_env = ttk.Frame(notebook)
        notebook.add(tab_env, text=".env nastavení")

        frm = ttk.Frame(tab_env, padding=14)
        frm.pack(fill="both", expand=True)
        frm.columnconfigure(1, weight=1)

        row = 0
        for name, label, wtype, opts in FIELDS:
            ttk.Label(frm, text=label).grid(row=row, column=0, sticky="w", pady=6)
            var = tk.StringVar()
            self.vars[name] = var

            if wtype == "entry":
                ent = ttk.Entry(frm, textvariable=var)
                ent.grid(row=row, column=1, sticky="we", padx=8)

            elif wtype == "entry_secret":
                wrap = ttk.Frame(frm); wrap.grid(row=row, column=1, sticky="we", padx=8)
                ent = ttk.Entry(wrap, textvariable=var, show="•"); ent.pack(side="left", fill="x", expand=True)
                def toggle_secret(entry_widget=ent, button_widget=None):
                    if entry_widget.cget("show") == "•":
                        entry_widget.config(show=""); button_widget and button_widget.config(text="Skrýt")
                    else:
                        entry_widget.config(show="•"); button_widget and button_widget.config(text="Zobrazit")
                btn = ttk.Button(wrap, text="Zobrazit", width=9)
                btn.config(command=lambda ew=ent, bw=btn: toggle_secret(entry_widget=ew, button_widget=bw))
                btn.pack(side="left", padx=6)

            elif wtype == "combo":
                cb = ttk.Combobox(frm, textvariable=var, state="readonly",
                                  values=opts.get("values", []))
                cb.grid(row=row, column=1, sticky="we", padx=8)
                if "default" in opts:
                    cb.set(opts["default"])

            elif wtype == "entry_browse":
                wrap = ttk.Frame(frm); wrap.grid(row=row, column=1, sticky="we", padx=8)
                ent = ttk.Entry(wrap, textvariable=var); ent.pack(side="left", fill="x", expand=True)
                def browse(v=var):
                    initial = v.get().strip() or "template.xlsx"
                    path = filedialog.askopenfilename(
                        initialfile=initial,
                        filetypes=[("Excel", "*.xlsx"), ("All files", "*.*")]
                    )
                    if path: v.set(path)
                ttk.Button(wrap, text="Procházet…", command=browse).pack(side="left", padx=6)

            elif wtype == "dir_browse":
                wrap = ttk.Frame(frm); wrap.grid(row=row, column=1, sticky="we", padx=8)
                ent = ttk.Entry(wrap, textvariable=var); ent.pack(side="left", fill="x", expand=True)
                def pick_dir(v=var):
                    path = filedialog.askdirectory()
                    if path: v.set(path)
                ttk.Button(wrap, text="Vybrat složku…", command=pick_dir).pack(side="left", padx=6)

            row += 1

        btns = ttk.Frame(frm); btns.grid(row=row, column=0, columnspan=2, sticky="we", pady=(14, 0))
        ttk.Button(btns, text="Obnovit výchozí", command=self.reset_defaults).pack(side="left")
        ttk.Button(btns, text="Uložit .env", command=self.save_env).pack(side="left", padx=6)
        ttk.Button(btns, text="Otevřít .env", command=lambda: os.startfile(ENV_PATH)).pack(side="left", padx=6)  # NEW
        ttk.Button(btns, text="Uložit & Spustit", command=self.save_and_run_interactive).pack(side="right")

        tab_log = ttk.Frame(notebook); notebook.add(tab_log, text="Logy")
        log_frame = ttk.Frame(tab_log, padding=10); log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=22, wrap="word"); self.log_text.pack(fill="both", expand=True, side="left")
        self.log_text.configure(state="disabled")
        scroll = ttk.Scrollbar(log_frame, orient="vertical", command=self.log_text.yview); scroll.pack(side="right", fill="y")
        self.log_text["yscrollcommand"] = scroll.set

        self.status = tk.StringVar(value=f"Cesta k .env: {ENV_PATH}")
        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill="x", padx=12, pady=6)

        tab_rules = ttk.Frame(notebook)
        notebook.add(tab_rules, text="Prompt rules")

        rules_frame = ttk.Frame(tab_rules, padding=10)
        rules_frame.pack(fill="both", expand=True)

        ttk.Label(rules_frame, text="PROMPT_RULES (multiline text saved to .env)").pack(anchor="w")

        wrap = ttk.Frame(rules_frame)
        wrap.pack(fill="both", expand=True, pady=(6, 0))

        self.prompt_text = tk.Text(wrap, height=18, wrap="word")  # NEW
        self.prompt_text.pack(side="left", fill="both", expand=True)

        rules_scroll = ttk.Scrollbar(wrap, orient="vertical", command=self.prompt_text.yview)
        rules_scroll.pack(side="right", fill="y")
        self.prompt_text["yscrollcommand"] = rules_scroll.set

    def _load_env(self):

        data = dotenv_values(ENV_PATH) if os.path.exists(ENV_PATH) else {}
        if not data and os.path.exists(EXAMPLE_PATH):
            data = dotenv_values(EXAMPLE_PATH)
        for name, _, _, opts in FIELDS:
            default = opts.get("default", "")
            self.vars[name].set((data.get(name) or default))

        rules_val = (data.get("PROMPT_RULES") or "")

        rules_val = rules_val.replace("\\n", "\n")

        if len(rules_val) >= 2 and rules_val[0] == rules_val[-1] == '"':
            rules_val = rules_val[1:-1]
        if hasattr(self, "prompt_text"):
            self.prompt_text.delete("1.0", "end")
            self.prompt_text.insert("1.0", rules_val)

    def reset_defaults(self):
        for name, _, _, opts in FIELDS:
            self.vars[name].set(opts.get("default", ""))
        if hasattr(self, "prompt_text"):
            self.prompt_text.delete("1.0", "end")

    def save_env(self) -> bool:
        existing_lines = []
        if os.path.exists(ENV_PATH):
            try:
                with open(ENV_PATH, "r", encoding="utf-8") as f:
                    existing_lines = f.read().splitlines()
            except Exception:
                existing_lines = []

        known_keys = {name for name, _, _, _ in FIELDS} | {"PROMPT_RULES"}
        preserved = []
        # Keep unknown keys from existing .env
        for line in existing_lines:
            striped = line.strip()
            if (not striped) or striped.startswith("#"):
                preserved.append(line); continue
            key = striped.split("=", 1)[0].strip()
            if key not in known_keys:
                preserved.append(line)

        if not existing_lines and os.path.exists(EXAMPLE_PATH):
            try:
                with open(EXAMPLE_PATH, "r", encoding="utf-8") as f:
                    for line in f.read().splitlines():
                        striped = line.strip()
                        if (not striped) or striped.startswith("#"):
                            preserved.append(line); continue
                        key = striped.split("=", 1)[0].strip()
                        if key not in known_keys:
                            preserved.append(line)
            except Exception:
                pass

        try:
            if "DATE_FROM" in self.vars:
                self.vars["DATE_FROM"].set(_coerce_date_to_iso(self.vars["DATE_FROM"].get()))
            if "DATE_TO" in self.vars:
                self.vars["DATE_TO"].set(_coerce_date_to_iso(self.vars["DATE_TO"].get()))
        except ValueError as e:
            messagebox.showerror("Chyba datumu", str(e)); self._log(f"[error] {e}\n"); return False

        new_lines = []
        for name, _, _, _ in FIELDS:
            val = self.vars[name].get().strip()
            if "\n" in val or '"' in val:
                val = val.replace('"', '\\"').replace("\n", "\\n")
                line = f'{name}="{val}"'
            else:
                line = f"{name}={val}"
            new_lines.append(line)
        rules_raw = self.prompt_text.get("1.0", "end-1c") if hasattr(self, "prompt_text") else ""
        rules_escaped = rules_raw.replace('"', '\\"').replace("\r\n", "\n").replace("\n", "\\n")
        new_lines.append(f'PROMPT_RULES="{rules_escaped}"')
        try:
            with open(ENV_PATH, "w", encoding="utf-8") as f:
                if preserved:
                    f.write("\n".join(preserved).rstrip() + "\n")
                if new_lines:
                    f.write("\n".join(new_lines) + "\n")
            self.status.set(f"Uloženo: {ENV_PATH}")
            self._log(f"[ok] .env uloženo → {ENV_PATH}\n")
            return True
        except Exception as e:
            messagebox.showerror("Chyba", f"Nepodařilo se uložit .env:\n{e}")
            self._log(f"[error] uložení .env selhalo: {e}\n")
            return False

    def save_and_run_interactive(self):
        if not self.save_env():
            return
        self._run_main(args=[], new_console=True)

    def _run_main(self, args, new_console=False):
        if self.proc and self.proc.poll() is None:
            messagebox.showwarning("Běží", "Aplikace již běží. Počkejte na dokončení.")
            return

        cmd = _cli_cmd() + args
        self._log(f"[run] {' '.join(cmd)}\n")

        def worker():
            try:
                self.proc = subprocess.Popen(
                    cmd, cwd=APP_DIR,
                    stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                    text=True, bufsize=1, universal_newlines=True,
                    encoding="utf-8", errors="replace",
                    env={**os.environ, "PYTHONIOENCODING": "utf-8"}
                )
                for line in self.proc.stdout:
                    self._log(line)
                rc = self.proc.wait()
                self._log(f"\n[exit] návratový kód: {rc}\n")
                if rc == 0:
                    messagebox.showinfo("Hotovo", "Aplikace úspěšně dokončila běh.")
                else:
                    messagebox.showwarning("Ukončeno s chybami", f"Proces skončil kódem {rc}")
            except Exception as e:
                self._log(f"[error] spuštění main.py selhalo: {e}\n")
                messagebox.showerror("Chyba", f"Nepodařilo se spustit main.py:\n{e}")

        threading.Thread(target=worker, daemon=True).start()

    def _log(self, msg: str):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

if __name__ == "__main__":
    if "--run-main" in sys.argv:
        from main import main as run_pipeline
        run_pipeline()
        sys.exit(0)
    app = EnvRunner()
    app.mainloop()






