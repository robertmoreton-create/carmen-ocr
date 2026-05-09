#!/usr/bin/env python3
"""Simple desktop app for HSBC statement OCR to Excel conversion."""

from __future__ import annotations

import json
import re
import sys
from datetime import date
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from hsbc_statement_to_excel import convert_statement

if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
    BUNDLE_DIR = Path(getattr(sys, "_MEIPASS", APP_DIR))
else:
    APP_DIR = Path(__file__).resolve().parent
    BUNDLE_DIR = APP_DIR

SETTINGS_PATH = APP_DIR / "app_settings.json"


def default_settings() -> dict:
    return {
        "input_folder": "{APP_DIR}/Statements/Incoming",
        "output_folder": "{APP_DIR}/Statements/Output",
        "template_workbook": "{BUNDLE_DIR}/Su - HSBC One #833 (PayMe).xlsx",
        "template_sheet": "26",
        "sheet_name_format": "%b-%Y",
    }


def load_settings() -> dict:
    if not SETTINGS_PATH.exists():
        settings = default_settings()
        save_settings(settings)
        return settings

    with SETTINGS_PATH.open("r", encoding="utf-8") as f:
        settings = json.load(f)

    defaults = default_settings()
    for key, value in defaults.items():
        settings.setdefault(key, value)
    return settings


def save_settings(settings: dict) -> None:
    SETTINGS_PATH.parent.mkdir(parents=True, exist_ok=True)
    with SETTINGS_PATH.open("w", encoding="utf-8") as f:
        json.dump(settings, f, indent=2)


def resolve_setting_path(raw: str) -> Path:
    value = (raw or "").strip().replace("\\", "/")
    value = value.replace("{APP_DIR}", str(APP_DIR))
    value = value.replace("{BUNDLE_DIR}", str(BUNDLE_DIR))
    return Path(value).expanduser()


def sanitize_sheet_name_for_file(value: str) -> str:
    cleaned = "".join(ch for ch in (value or "").strip() if ord(ch) >= 32)
    cleaned = re.sub(r'[<>:"/\\\\|?*]', "-", cleaned)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    return cleaned[:64] or date.today().strftime("%b-%Y")


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("HSBC Statement OCR")
        self.geometry("900x520")

        self.settings = load_settings()

        self.input_var = tk.StringVar(value=self.settings["input_folder"])
        self.output_var = tk.StringVar(value=self.settings["output_folder"])
        self.template_var = tk.StringVar(value=self.settings["template_workbook"])
        self.template_sheet_var = tk.StringVar(value=self.settings["template_sheet"])
        self.sheet_name_var = tk.StringVar(value=date.today().strftime(self.settings["sheet_name_format"]))

        self._build_ui()
        self.ensure_runtime_folders()
        self.refresh_pdf_list()

    def _build_ui(self) -> None:
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        ttk.Label(root, text="Input Folder").grid(row=0, column=0, sticky="w")
        ttk.Entry(root, textvariable=self.input_var, width=90).grid(row=0, column=1, sticky="ew", padx=6)
        ttk.Button(root, text="Browse", command=self.pick_input_folder).grid(row=0, column=2)

        ttk.Label(root, text="Output Folder").grid(row=1, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(root, textvariable=self.output_var, width=90).grid(row=1, column=1, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(root, text="Browse", command=self.pick_output_folder).grid(row=1, column=2, pady=(8, 0))

        ttk.Label(root, text="Template Workbook").grid(row=2, column=0, sticky="w", pady=(8, 0))
        ttk.Entry(root, textvariable=self.template_var, width=90).grid(row=2, column=1, sticky="ew", padx=6, pady=(8, 0))
        ttk.Button(root, text="Browse", command=self.pick_template_file).grid(row=2, column=2, pady=(8, 0))

        settings_row = ttk.Frame(root)
        settings_row.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 0))
        ttk.Label(settings_row, text="Template Sheet").pack(side="left")
        ttk.Entry(settings_row, textvariable=self.template_sheet_var, width=10).pack(side="left", padx=(8, 18))
        ttk.Label(settings_row, text="Output Sheet Name").pack(side="left")
        ttk.Entry(settings_row, textvariable=self.sheet_name_var, width=18).pack(side="left", padx=(8, 10))
        ttk.Button(settings_row, text="Save Settings", command=self.save_current_settings).pack(side="left")

        ttk.Separator(root).grid(row=4, column=0, columnspan=3, sticky="ew", pady=12)

        list_header = ttk.Frame(root)
        list_header.grid(row=5, column=0, columnspan=3, sticky="ew")
        ttk.Label(list_header, text="Statement PDFs in Input Folder").pack(side="left")
        ttk.Button(list_header, text="Refresh", command=self.refresh_pdf_list).pack(side="right")

        self.pdf_list = tk.Listbox(root, height=12)
        self.pdf_list.grid(row=6, column=0, columnspan=3, sticky="nsew", pady=(8, 10))

        buttons = ttk.Frame(root)
        buttons.grid(row=7, column=0, columnspan=3, sticky="ew")
        ttk.Button(buttons, text="Generate Excel From Selected PDF", command=self.process_selected).pack(side="left")

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(root, textvariable=self.status_var).grid(row=8, column=0, columnspan=3, sticky="w", pady=(10, 0))

        root.columnconfigure(1, weight=1)
        root.rowconfigure(6, weight=1)

    def pick_input_folder(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.input_var.get() or str(APP_DIR))
        if chosen:
            self.input_var.set(chosen)
            self.refresh_pdf_list()

    def pick_output_folder(self) -> None:
        chosen = filedialog.askdirectory(initialdir=self.output_var.get() or str(APP_DIR))
        if chosen:
            self.output_var.set(chosen)

    def pick_template_file(self) -> None:
        chosen = filedialog.askopenfilename(
            initialdir=str(Path(self.template_var.get()).parent if self.template_var.get() else APP_DIR),
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if chosen:
            self.template_var.set(chosen)

    def save_current_settings(self) -> None:
        settings = {
            "input_folder": self.input_var.get().strip(),
            "output_folder": self.output_var.get().strip(),
            "template_workbook": self.template_var.get().strip(),
            "template_sheet": self.template_sheet_var.get().strip() or "26",
            "sheet_name_format": self.settings.get("sheet_name_format", "%b-%Y"),
        }
        save_settings(settings)
        self.settings = settings
        self.ensure_runtime_folders()
        self.status_var.set(f"Saved settings to {SETTINGS_PATH}")

    def ensure_runtime_folders(self) -> None:
        for raw_path in (self.input_var.get(), self.output_var.get()):
            resolve_setting_path(raw_path).mkdir(parents=True, exist_ok=True)

    def refresh_pdf_list(self) -> None:
        folder = resolve_setting_path(self.input_var.get())
        self.pdf_list.delete(0, tk.END)
        if not folder.exists():
            self.status_var.set(f"Input folder does not exist: {folder}")
            return

        files = sorted(folder.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        for pdf in files:
            self.pdf_list.insert(tk.END, pdf.name)

        self.status_var.set(f"Found {len(files)} PDF file(s)")

    def process_selected(self) -> None:
        selection = self.pdf_list.curselection()
        if not selection:
            messagebox.showwarning("No PDF selected", "Select a statement PDF first.")
            return

        input_folder = resolve_setting_path(self.input_var.get())
        output_folder = resolve_setting_path(self.output_var.get())
        template_path = resolve_setting_path(self.template_var.get())
        template_sheet = self.template_sheet_var.get().strip() or "26"
        sheet_name = self.sheet_name_var.get().strip() or date.today().strftime("%b-%Y")
        safe_sheet_name = sanitize_sheet_name_for_file(sheet_name)

        pdf_name = self.pdf_list.get(selection[0])
        pdf_path = input_folder / pdf_name

        output_folder.mkdir(parents=True, exist_ok=True)
        output_file = output_folder / f"HSBC_{safe_sheet_name}.xlsx"

        self.status_var.set("Processing PDF with OCR...")
        self.update_idletasks()

        try:
            result = convert_statement(
                pdf_path=pdf_path,
                template_path=template_path,
                output_path=output_file,
                template_sheet_name=template_sheet,
                sheet_name=sheet_name,
                order_mode="strict_pdf",
            )
        except Exception as exc:
            messagebox.showerror("Processing failed", str(exc))
            self.status_var.set("Failed")
            return

        self.status_var.set(f"Done: {result['entries_written']} entries -> {output_file}")
        messagebox.showinfo(
            "Completed",
            f"Sheet: {result['sheet_name']}\n"
            f"Entries: {result['entries_written']}\n"
            f"Saved: {output_file}",
        )


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    if sys.platform.startswith("win"):
        # Better DPI behavior on Windows displays.
        try:
            from ctypes import windll

            windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
    main()
