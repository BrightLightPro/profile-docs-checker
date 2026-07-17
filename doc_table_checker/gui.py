from __future__ import annotations

import os
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from .core import create_corrected_word_copy, run_validation


class App(tk.Tk):
    """Light control-panel interface with a compact on-screen validation report."""

    BG = "#EEF4F6"
    PANEL = "#FFFFFF"
    PANEL_ALT = "#E8F1F4"
    PANEL_HOVER = "#DCECEF"
    BORDER = "#BFD1D8"
    ACCENT = "#00AFC3"
    ACCENT_DARK = "#007A8A"
    TEXT = "#20343C"
    MUTED = "#667A83"
    WARNING = "#AE7200"
    WARNING_BG = "#FFF4D8"
    ERROR = "#C94855"
    ERROR_BG = "#FCE4E7"
    SUCCESS = "#247F5C"
    SUCCESS_BG = "#E2F4EB"

    def __init__(self):
        super().__init__()
        self.title("Profile Docs Checker")
        self.geometry("1240x790")
        self.minsize(1080, 700)
        self.configure(bg=self.BG)

        self.pdf_folder_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.word_file_var = tk.StringVar()
        self.ausgabe_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.mode_text_var = tk.StringVar(value="Select an Excel file and at least one document source.")
        self.current_step_var = tk.StringVar(value="Ready.")
        self.result_title_var = tk.StringVar(value="AWAITING VALIDATION")
        self.result_detail_var = tk.StringVar(value="Run the checker to see the complete concise result here.")
        self._last_summary: dict = {}
        self._last_run_succeeded = False

        self._configure_styles()
        self._build_ui()
        self._refresh_mode_status()
        self.after_idle(self._maximize_or_fit)

        self.pdf_folder_var.trace_add("write", lambda *_: self._refresh_mode_status())
        self.word_file_var.trace_add("write", lambda *_: self._refresh_mode_status())

    def _maximize_or_fit(self):
        try:
            if os.name == "nt":
                self.state("zoomed")
                return
        except tk.TclError:
            pass
        self.update_idletasks()
        screen_w = self.winfo_screenwidth()
        screen_h = self.winfo_screenheight()
        width = min(1320, max(1080, screen_w - 50))
        height = min(860, max(700, screen_h - 70))
        x = max(0, (screen_w - width) // 2)
        y = max(0, (screen_h - height) // 2)
        self.geometry(f"{width}x{height}+{x}+{y}")

    # ------------------------------------------------------------------ styling
    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Accent.TButton", background=self.ACCENT, foreground="#FFFFFF",
            bordercolor=self.ACCENT, lightcolor=self.ACCENT, darkcolor=self.ACCENT_DARK,
            relief="flat", padding=(18, 10), font=("Segoe UI Semibold", 10),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", "#D7E3E7"), ("active", "#18C0D1")],
            foreground=[("disabled", self.MUTED), ("active", "#FFFFFF")],
        )
        style.configure(
            "Ghost.TButton", background=self.PANEL_ALT, foreground=self.TEXT,
            bordercolor=self.BORDER, lightcolor=self.PANEL_ALT, darkcolor=self.PANEL_ALT,
            relief="flat", padding=(12, 8), font=("Segoe UI", 9),
        )
        style.map(
            "Ghost.TButton",
            background=[("active", self.PANEL_HOVER), ("disabled", self.PANEL)],
            foreground=[("disabled", self.MUTED)],
        )
        style.configure(
            "Compact.TButton", background=self.PANEL_ALT, foreground=self.TEXT,
            bordercolor=self.BORDER, lightcolor=self.PANEL_ALT, darkcolor=self.PANEL_ALT,
            relief="flat", padding=(8, 7), font=("Segoe UI", 8),
        )
        style.map("Compact.TButton", background=[("active", self.PANEL_HOVER)])
        style.configure(
            "Light.TCombobox", fieldbackground="#F8FBFC", background="#F8FBFC",
            foreground=self.TEXT, arrowcolor=self.ACCENT, bordercolor=self.BORDER,
            lightcolor=self.BORDER, darkcolor=self.BORDER, padding=5,
        )
        style.map(
            "Light.TCombobox",
            fieldbackground=[("readonly", "#F8FBFC")],
            foreground=[("readonly", self.TEXT)],
            selectbackground=[("readonly", "#F8FBFC")],
            selectforeground=[("readonly", self.TEXT)],
        )
        style.configure(
            "Control.Horizontal.TProgressbar", troughcolor=self.PANEL,
            background=self.ACCENT, bordercolor=self.BORDER, lightcolor=self.ACCENT,
            darkcolor=self.ACCENT_DARK, thickness=5,
        )
        style.configure(
            "Results.Treeview", background="#F8FBFC", fieldbackground="#F8FBFC",
            foreground=self.TEXT, rowheight=38, bordercolor=self.BORDER,
            lightcolor=self.BORDER, darkcolor=self.BORDER, font=("Segoe UI", 9),
        )
        style.configure(
            "Results.Treeview.Heading", background=self.PANEL_ALT, foreground=self.TEXT,
            relief="flat", font=("Segoe UI Semibold", 9), padding=(8, 7),
        )
        style.map("Results.Treeview", background=[("selected", "#CDECF1")], foreground=[("selected", self.TEXT)])

    # ------------------------------------------------------------------ layout
    def _build_ui(self):
        self._build_header()
        body = tk.Frame(self, bg=self.BG)
        body.pack(fill="both", expand=True, padx=24, pady=(18, 20))
        body.grid_columnconfigure(0, weight=6, minsize=520)
        body.grid_columnconfigure(1, weight=6, minsize=470)
        body.grid_rowconfigure(0, weight=1)

        left = tk.Frame(body, bg=self.PANEL, highlightthickness=1, highlightbackground=self.BORDER)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        left.grid_columnconfigure(0, weight=1)

        right = tk.Frame(body, bg=self.PANEL, highlightthickness=1, highlightbackground=self.BORDER)
        right.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        right.grid_columnconfigure(0, weight=1)
        right.grid_rowconfigure(4, weight=1)

        self._build_configuration(left)
        self._build_results_panel(right)

    def _build_header(self):
        header = tk.Canvas(self, height=112, bg=self.BG, highlightthickness=0)
        header.pack(fill="x")
        header.bind("<Configure>", lambda event: self._draw_header(header, event.width))

    def _draw_header(self, canvas: tk.Canvas, width: int):
        canvas.delete("all")
        canvas.create_polygon(
            0, 0, width, 0, width, 84, width - 38, 112, 38, 112, 0, 84,
            fill=self.PANEL, outline=self.BORDER,
        )
        canvas.create_line(42, 111, width - 42, 111, fill=self.ACCENT, width=2)
        canvas.create_polygon(0, 0, 210, 0, 180, 12, 0, 12, fill=self.ACCENT_DARK, outline="")
        canvas.create_text(44, 43, anchor="w", text="PROFILE // DOCS CHECKER", fill=self.TEXT, font=("Segoe UI Semibold", 19))
        canvas.create_text(45, 73, anchor="w", text="DOCUMENT CONSISTENCY CONTROL", fill=self.ACCENT, font=("Consolas", 9, "bold"))
        canvas.create_text(width - 45, 50, anchor="e", text="PDF  •  EXCEL  •  WORD", fill=self.MUTED, font=("Segoe UI", 10))
        canvas.create_text(width - 45, 74, anchor="e", text="v1.5.0", fill=self.ACCENT, font=("Consolas", 9))

    def _build_configuration(self, parent: tk.Frame):
        title_row = tk.Frame(parent, bg=self.PANEL)
        title_row.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(title_row, text="VALIDATION INPUTS", bg=self.PANEL, fg=self.TEXT, font=("Segoe UI Semibold", 12)).pack(side="left")
        tk.Label(title_row, text="CONTROL", bg=self.PANEL_ALT, fg=self.ACCENT_DARK, padx=10, pady=4, font=("Consolas", 8, "bold")).pack(side="right")
        tk.Frame(parent, height=1, bg=self.BORDER).pack(fill="x", padx=20, pady=(0, 14))

        fields = tk.Frame(parent, bg=self.PANEL)
        fields.pack(fill="both", expand=True, padx=20)
        fields.grid_columnconfigure(0, weight=1)

        self._path_field(fields, 0, "EXCEL MASTER LIST  •  REQUIRED", self.excel_file_var, self.choose_excel_file, "Fixed German master template (.xlsx / .xlsm)")

        sheet_row = tk.Frame(fields, bg=self.PANEL)
        sheet_row.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        sheet_row.grid_columnconfigure(0, weight=1)
        tk.Label(sheet_row, text="EXCEL SHEET", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.sheet_var, state="readonly", style="Light.TCombobox")
        self.sheet_combo.grid(row=1, column=0, sticky="ew")
        ttk.Button(sheet_row, text="RELOAD", style="Compact.TButton", command=self.load_sheets).grid(row=1, column=1, padx=(8, 0))

        self._path_field(
            fields, 2, "WORD CHANGE NOTICE  •  OPTIONAL", self.word_file_var, self.choose_word_file,
            "Selecting a Word file automatically enables Word-vs-Excel and Blatt/page checks.",
            clear_command=lambda: self.word_file_var.set(""),
        )
        self._path_field(
            fields, 3, "PDF DOCUMENT FOLDER  •  OPTIONAL", self.pdf_folder_var, self.choose_pdf_folder,
            "Required for PDF metadata checks and actual page counts. Metadata is read from page 2.",
            clear_command=lambda: self.pdf_folder_var.set(""),
        )

        issue_row = tk.Frame(fields, bg=self.PANEL)
        issue_row.grid(row=4, column=0, sticky="ew", pady=(0, 14))
        issue_row.grid_columnconfigure(0, weight=1)
        tk.Label(issue_row, text="EXPECTED AUSGABE", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.ausgabe_entry = self._light_entry(issue_row, self.ausgabe_var)
        self.ausgabe_entry.grid(row=1, column=0, sticky="ew")
        tk.Label(issue_row, text="Month and year, e.g. 07.2026. Required only when PDFs are selected.", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8)).grid(row=2, column=0, sticky="w", pady=(5, 0))

        self._path_field(fields, 5, "OUTPUT REPORT", self.output_file_var, self.choose_output_file, "Two-sheet Excel validation report (.xlsx)", browse_text="SAVE AS")

        actions = tk.Frame(parent, bg=self.PANEL)
        actions.pack(fill="x", padx=20, pady=(0, 18))
        self.run_button = ttk.Button(actions, text="RUN VALIDATION", style="Accent.TButton", command=self.run)
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(actions, text="OPEN EXCEL REPORT", style="Ghost.TButton", command=self.open_report, state="disabled")
        self.open_button.pack(side="left", padx=(9, 0))

    def _build_results_panel(self, parent: tk.Frame):
        top = tk.Frame(parent, bg=self.PANEL)
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        tk.Label(top, text="VALIDATION RESULT", bg=self.PANEL, fg=self.TEXT, font=("Segoe UI Semibold", 11)).pack(side="left")
        self.activity_label = tk.Label(top, text="IDLE", bg=self.PANEL_ALT, fg=self.MUTED, padx=9, pady=3, font=("Consolas", 8, "bold"))
        self.activity_label.pack(side="right")

        self.status_card = tk.Frame(parent, bg=self.PANEL_ALT, highlightthickness=1, highlightbackground=self.BORDER)
        self.status_card.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 10))
        self.status_card.grid_columnconfigure(0, weight=1)
        self.status_title_label = tk.Label(self.status_card, textvariable=self.result_title_var, bg=self.PANEL_ALT, fg=self.ACCENT_DARK, font=("Segoe UI Semibold", 15), anchor="w")
        self.status_title_label.grid(row=0, column=0, sticky="ew", padx=14, pady=(12, 2))
        self.status_detail_label = tk.Label(self.status_card, textvariable=self.result_detail_var, bg=self.PANEL_ALT, fg=self.TEXT, font=("Segoe UI", 9), anchor="w", justify="left", wraplength=520)
        self.status_detail_label.grid(row=1, column=0, sticky="ew", padx=14, pady=(0, 12))

        self.progress = ttk.Progressbar(parent, mode="determinate", value=0, maximum=100, style="Control.Horizontal.TProgressbar")
        self.progress.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 8))

        mode_strip = tk.Frame(parent, bg=self.PANEL_ALT, highlightthickness=1, highlightbackground=self.BORDER)
        mode_strip.grid(row=3, column=0, sticky="ew", padx=16, pady=(0, 10))
        tk.Label(mode_strip, text="MODE", bg=self.PANEL_ALT, fg=self.ACCENT, font=("Consolas", 8, "bold")).pack(side="left", padx=(10, 8), pady=7)
        tk.Label(mode_strip, textvariable=self.mode_text_var, bg=self.PANEL_ALT, fg=self.TEXT, font=("Segoe UI", 8), wraplength=470, justify="left").pack(side="left", fill="x", expand=True, padx=(0, 10), pady=7)

        table_frame = tk.Frame(parent, bg="#F8FBFC", highlightthickness=1, highlightbackground=self.BORDER)
        table_frame.grid(row=4, column=0, sticky="nsew", padx=16, pady=(0, 10))
        table_frame.grid_rowconfigure(0, weight=1)
        table_frame.grid_columnconfigure(0, weight=1)
        self.result_tree = ttk.Treeview(table_frame, columns=("id", "issue"), show="headings", style="Results.Treeview", selectmode="browse")
        self.result_tree.heading("id", text="ID-NR.")
        self.result_tree.heading("issue", text="INCORRECT OR UNVERIFIED INFORMATION")
        self.result_tree.column("id", width=125, minwidth=90, stretch=False, anchor="w")
        self.result_tree.column("issue", width=410, minwidth=260, stretch=True, anchor="w")
        self.result_tree.grid(row=0, column=0, sticky="nsew")
        scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.result_tree.yview)
        scroll.grid(row=0, column=1, sticky="ns")
        self.result_tree.configure(yscrollcommand=scroll.set)
        self.result_tree.tag_configure("error", background=self.ERROR_BG, foreground="#7D2630")
        self.result_tree.tag_configure("warning", background=self.WARNING_BG, foreground="#795000")
        self.result_tree.tag_configure("success", background=self.SUCCESS_BG, foreground=self.SUCCESS)
        self.result_tree.tag_configure("neutral", background="#F8FBFC", foreground=self.MUTED)
        self._set_placeholder_result()

        footer = tk.Frame(parent, bg=self.PANEL)
        footer.grid(row=5, column=0, sticky="ew", padx=16, pady=(0, 16))
        footer.grid_columnconfigure(0, weight=1)
        tk.Label(footer, textvariable=self.current_step_var, bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8), anchor="w", justify="left", wraplength=400).grid(row=0, column=0, sticky="ew")
        self.correct_button = ttk.Button(footer, text="CREATE CORRECTED WORD COPY", style="Ghost.TButton", command=self.correct_word, state="disabled")
        self.correct_button.grid(row=0, column=1, padx=(10, 0))

    def _set_placeholder_result(self):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.result_tree.insert("", "end", values=("—", "No validation result yet."), tags=("neutral",))

    def _light_entry(self, parent, variable: tk.StringVar) -> tk.Entry:
        return tk.Entry(
            parent, textvariable=variable, bg="#F8FBFC", fg=self.TEXT,
            insertbackground=self.ACCENT, selectbackground=self.ACCENT_DARK,
            selectforeground="#FFFFFF", relief="flat", highlightthickness=1,
            highlightbackground=self.BORDER, highlightcolor=self.ACCENT,
            disabledbackground="#EEF3F5", disabledforeground=self.MUTED,
            font=("Segoe UI", 9),
        )

    def _path_field(self, parent, row, label, variable, browse_command, file_hint, clear_command=None, browse_text="BROWSE"):
        frame = tk.Frame(parent, bg=self.PANEL)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 14))
        frame.grid_columnconfigure(0, weight=1)
        tk.Label(frame, text=label, bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self._light_entry(frame, variable).grid(row=1, column=0, sticky="ew")
        ttk.Button(frame, text=browse_text, style="Compact.TButton", command=browse_command).grid(row=1, column=1, padx=(8, 0))
        if clear_command:
            ttk.Button(frame, text="CLEAR", style="Compact.TButton", command=clear_command).grid(row=1, column=2, padx=(6, 0))
        tk.Label(frame, text=file_hint, bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8), wraplength=450, justify="left").grid(row=2, column=0, columnspan=3, sticky="w", pady=(5, 0))

    # ------------------------------------------------------------------ file actions
    def choose_pdf_folder(self):
        path = filedialog.askdirectory(title="Choose folder containing PDFs")
        if path:
            self.pdf_folder_var.set(path)
            self._suggest_output(Path(path))

    def choose_excel_file(self):
        path = filedialog.askopenfilename(title="Choose Excel master list", filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.excel_file_var.set(path)
            self.load_sheets()
            self._suggest_output(Path(path).parent)

    def choose_word_file(self):
        path = filedialog.askopenfilename(title="Choose Word change notice", filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
        if path:
            self.word_file_var.set(path)
            self._suggest_output(Path(path).parent)

    def choose_output_file(self):
        initial = self.output_file_var.get().strip() or "profile_docs_validation_report.xlsx"
        parent = Path(initial).parent
        path = filedialog.asksaveasfilename(
            title="Save validation report as", defaultextension=".xlsx",
            initialdir=str(parent) if parent.exists() else None,
            initialfile=Path(initial).name, filetypes=[("Excel report", "*.xlsx")],
        )
        if path:
            self.output_file_var.set(path)

    def _suggest_output(self, base: Path):
        if not self.output_file_var.get().strip():
            folder = base if base.is_dir() else base.parent
            self.output_file_var.set(str(folder / "profile_docs_validation_report.xlsx"))

    def load_sheets(self):
        excel = self.excel_file_var.get().strip()
        if not excel:
            return
        try:
            wb = load_workbook(excel, read_only=True, data_only=True)
            sheets = wb.sheetnames
            wb.close()
            self.sheet_combo["values"] = sheets
            if sheets and self.sheet_var.get() not in sheets:
                self.sheet_var.set(sheets[0])
            self.current_step_var.set("Excel sheets loaded: " + ", ".join(sheets))
        except Exception as exc:
            messagebox.showerror("Could not load Excel sheets", str(exc))

    # ------------------------------------------------------------------ state
    def _refresh_mode_status(self):
        pdf = bool(self.pdf_folder_var.get().strip())
        word = bool(self.word_file_var.get().strip())
        if pdf and word:
            text = "FULL CONTROL  •  PDF ↔ Excel  •  Word ↔ Excel  •  Word pages ↔ actual PDF pages"
        elif pdf:
            text = "PDF CONTROL  •  PDF metadata ↔ Excel  •  Ausgabe validation"
        elif word:
            text = "WORD CONTROL  •  Word ↔ Excel  •  page count remains unverified without PDFs"
        else:
            text = "Select an Excel file and at least one document source."
        self.mode_text_var.set(text)
        try:
            self.ausgabe_entry.configure(state="normal" if pdf else "disabled")
        except AttributeError:
            pass

    def _collect_args(self):
        pdf_folder = self.pdf_folder_var.get().strip() or None
        excel = self.excel_file_var.get().strip()
        word_file = self.word_file_var.get().strip() or None
        output = self.output_file_var.get().strip()
        ausgabe = self.ausgabe_var.get().strip() or None
        if not excel:
            raise ValueError("Choose the Excel master list.")
        if not pdf_folder and not word_file:
            raise ValueError("Choose a PDF folder, a Word file, or both.")
        if pdf_folder and not ausgabe:
            raise ValueError("Enter the expected Ausgabe when a PDF folder is selected.")
        if not output:
            raise ValueError("Choose an output report path.")
        return pdf_folder, excel, word_file, output, ausgabe

    # ------------------------------------------------------------------ validation
    def run(self):
        try:
            args = self._collect_args()
        except Exception as exc:
            messagebox.showerror("Missing or invalid input", str(exc))
            return
        pdf_folder, _excel, word_file, _output, _ausgabe = args
        if word_file and not pdf_folder:
            proceed = messagebox.askyesno(
                "PDF page count unavailable",
                "No PDF folder was selected.\n\nThe Word table will be compared with Excel, but Blatt / sheets cannot be verified against the true PDF page count. Continue?",
                icon="warning",
            )
            if not proceed:
                return

        self._last_run_succeeded = False
        self._last_summary = {}
        self.open_button.configure(state="disabled")
        self.correct_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.activity_label.configure(text="RUNNING", fg=self.ACCENT)
        self.result_title_var.set("VALIDATION IN PROGRESS")
        self.result_detail_var.set("The concise result will appear below as soon as all selected checks finish.")
        self.status_card.configure(bg=self.PANEL_ALT, highlightbackground=self.BORDER)
        self.status_title_label.configure(bg=self.PANEL_ALT, fg=self.ACCENT_DARK)
        self.status_detail_label.configure(bg=self.PANEL_ALT)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.result_tree.insert("", "end", values=("…", "Checking documents…"), tags=("neutral",))
        self.current_step_var.set("Starting validation sequence…")
        threading.Thread(target=self._run_thread, args=args, daemon=True).start()

    def _run_thread(self, pdf_folder, excel, word_file, output, ausgabe):
        summary_holder: dict = {}
        try:
            def progress(message: str):
                self.after(0, lambda m=message: self.current_step_var.set(m))

            def receive_summary(summary: dict):
                summary_holder.update(summary)

            run_validation(
                pdf_folder=pdf_folder, excel_path=excel, output_path=output,
                expected_ausgabe=ausgabe, sheet_name=self.sheet_var.get().strip() or None,
                page_number=2, word_docx_path=word_file,
                progress_callback=progress, summary_callback=receive_summary,
            )
            self.after(0, lambda: self._finish_success(output, summary_holder))
        except Exception as exc:
            self.after(0, lambda: self._finish_error(str(exc)))

    def _finish_success(self, output: str, summary: dict):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=100)
        self.run_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.activity_label.configure(text="COMPLETE", fg=self.SUCCESS)
        self.current_step_var.set(f"Excel report saved: {output}")
        self._last_run_succeeded = True
        self._last_summary = summary
        self._render_summary(summary)
        if self.word_file_var.get().strip():
            self.correct_button.configure(state="normal")

    def _render_summary(self, summary: dict):
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        issues = summary.get("issues", [])
        checked = int(summary.get("checked_entries", 0))
        if summary.get("all_correct", False):
            self.result_title_var.set("ALL ENTRIES ARE CORRECT")
            self.result_detail_var.set(f"All selected checks passed for {checked} Excel entr{'y' if checked == 1 else 'ies'}.")
            self.status_card.configure(bg=self.SUCCESS_BG, highlightbackground="#A9D8C1")
            self.status_title_label.configure(bg=self.SUCCESS_BG, fg=self.SUCCESS)
            self.status_detail_label.configure(bg=self.SUCCESS_BG)
            self.result_tree.insert("", "end", values=("✓", "All selected values are correct."), tags=("success",))
            return

        errors = int(summary.get("error_count", 0))
        warnings = int(summary.get("warning_count", 0))
        self.result_title_var.set(f"{len(issues)} ITEM{'S' if len(issues) != 1 else ''} NEED ATTENTION")
        detail = f"{errors} error{'s' if errors != 1 else ''}"
        if warnings:
            detail += f" and {warnings} warning{'s' if warnings != 1 else ''}"
        detail += ". Incorrect or unverified values are highlighted below."
        self.result_detail_var.set(detail)
        self.status_card.configure(bg=self.ERROR_BG if errors else self.WARNING_BG, highlightbackground="#E4B8BE" if errors else "#E5CB88")
        self.status_title_label.configure(bg=self.ERROR_BG if errors else self.WARNING_BG, fg=self.ERROR if errors else self.WARNING)
        self.status_detail_label.configure(bg=self.ERROR_BG if errors else self.WARNING_BG)
        for issue in issues:
            severity = issue.get("severity", "ERROR")
            dok_id = issue.get("dok_id") or "—"
            text = issue.get("issue", "")
            source = issue.get("source", "")
            if source:
                text = f"{text}  [{source}]"
            self.result_tree.insert("", "end", values=(dok_id, text), tags=("warning" if severity == "WARNING" else "error",))

    def _finish_error(self, error_text: str):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.run_button.configure(state="normal")
        self.activity_label.configure(text="ERROR", fg=self.ERROR)
        self.result_title_var.set("VALIDATION FAILED")
        self.result_detail_var.set(error_text)
        self.status_card.configure(bg=self.ERROR_BG, highlightbackground="#E4B8BE")
        self.status_title_label.configure(bg=self.ERROR_BG, fg=self.ERROR)
        self.status_detail_label.configure(bg=self.ERROR_BG)
        self.current_step_var.set("The validation did not finish.")
        for item in self.result_tree.get_children():
            self.result_tree.delete(item)
        self.result_tree.insert("", "end", values=("—", error_text), tags=("error",))
        messagebox.showerror("Validation failed", error_text)

    # ------------------------------------------------------------------ correction
    def correct_word(self):
        if not self._last_run_succeeded or not self.word_file_var.get().strip():
            messagebox.showerror("No Word result", "Run validation with a Word file first.")
            return
        source = Path(self.word_file_var.get().strip())
        suggested = source.with_name(source.stem + "_corrected.docx")
        output = filedialog.asksaveasfilename(
            title="Save corrected Word copy", defaultextension=".docx",
            initialdir=str(suggested.parent), initialfile=suggested.name,
            filetypes=[("Word document", "*.docx")],
        )
        if not output:
            return
        proceed = messagebox.askyesno(
            "Create corrected Word copy",
            "A new Word file will be created; the original will remain unchanged.\n\n"
            "Matched rows will receive exact Excel values. Blatt / sheets will use actual PDF page counts when PDFs are available. Continue?",
            icon="question",
        )
        if not proceed:
            return
        self.correct_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.activity_label.configure(text="CORRECTING", fg=self.ACCENT)
        self.current_step_var.set("Creating corrected Word copy…")
        threading.Thread(target=self._correct_word_thread, args=(output,), daemon=True).start()

    def _correct_word_thread(self, output: str):
        try:
            def progress(message: str):
                self.after(0, lambda m=message: self.current_step_var.set(m))
            result = create_corrected_word_copy(
                word_docx_path=self.word_file_var.get().strip(),
                excel_path=self.excel_file_var.get().strip(),
                output_path=output,
                sheet_name=self.sheet_var.get().strip() or None,
                pdf_folder=self.pdf_folder_var.get().strip() or None,
                progress_callback=progress,
            )
            self.after(0, lambda: self._finish_correction(result))
        except Exception as exc:
            self.after(0, lambda: self._finish_correction_error(str(exc)))

    def _finish_correction(self, result):
        self.run_button.configure(state="normal")
        self.correct_button.configure(state="normal")
        self.activity_label.configure(text="COMPLETE", fg=self.SUCCESS)
        self.current_step_var.set(f"Corrected Word copy saved: {result.output_path}")
        warning_text = ""
        if result.warnings:
            warning_text = "\n\nWarnings:\n- " + "\n- ".join(result.warnings)
        messagebox.showinfo(
            "Corrected Word copy created",
            f"Saved to:\n{result.output_path}\n\nChanged cells: {result.changed_cells}\nCorrected rows: {result.corrected_rows}\nSkipped rows: {result.skipped_rows}{warning_text}",
        )

    def _finish_correction_error(self, error_text: str):
        self.run_button.configure(state="normal")
        self.correct_button.configure(state="normal")
        self.activity_label.configure(text="ERROR", fg=self.ERROR)
        self.current_step_var.set("Could not create the corrected Word copy.")
        messagebox.showerror("Word correction failed", error_text)

    # ------------------------------------------------------------------ report
    def open_report(self):
        path = self.output_file_var.get().strip()
        if not path or not Path(path).exists():
            messagebox.showerror("Report not found", "The report file does not exist yet.")
            return
        try:
            if os.name == "nt":
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                webbrowser.open(Path(path).resolve().as_uri())
        except Exception as exc:
            messagebox.showerror("Could not open report", str(exc))


def main():
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
