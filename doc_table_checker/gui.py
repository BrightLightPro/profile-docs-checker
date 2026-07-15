from __future__ import annotations

import os
import threading
import webbrowser
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from .core import run_validation, summarize_results


class App(tk.Tk):
    """Dark desktop interface for PDF/Excel/Word profile validation."""

    BG = "#0A0E11"
    PANEL = "#11171B"
    PANEL_ALT = "#151D22"
    PANEL_HOVER = "#1A252B"
    BORDER = "#27363D"
    ACCENT = "#00C7D9"
    ACCENT_DARK = "#087C87"
    TEXT = "#EAF7F8"
    MUTED = "#83969C"
    WARNING = "#EAB84D"
    ERROR = "#FF6470"
    SUCCESS = "#5FD39A"

    def __init__(self):
        super().__init__()
        self.title("Profile Docs Checker")
        self.geometry("1120x800")
        self.minsize(1000, 700)
        self.configure(bg=self.BG)

        self.pdf_folder_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.word_file_var = tk.StringVar()
        self.ausgabe_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.mode_text_var = tk.StringVar(value="Select an Excel file and at least one document source.")

        self._configure_styles()
        self._build_ui()
        self._refresh_mode_status()

        self.pdf_folder_var.trace_add("write", lambda *_: self._refresh_mode_status())
        self.word_file_var.trace_add("write", lambda *_: self._refresh_mode_status())

    # ------------------------------------------------------------------ styling
    def _configure_styles(self):
        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure("Dark.TFrame", background=self.BG)
        style.configure("Panel.TFrame", background=self.PANEL)
        style.configure("PanelAlt.TFrame", background=self.PANEL_ALT)
        style.configure(
            "Section.TLabel",
            background=self.PANEL,
            foreground=self.TEXT,
            font=("Segoe UI Semibold", 11),
        )
        style.configure(
            "Field.TLabel",
            background=self.PANEL,
            foreground=self.MUTED,
            font=("Segoe UI", 9),
        )
        style.configure(
            "Hint.TLabel",
            background=self.PANEL,
            foreground=self.MUTED,
            font=("Segoe UI", 8),
        )
        style.configure(
            "Status.TLabel",
            background=self.PANEL_ALT,
            foreground=self.ACCENT,
            font=("Segoe UI Semibold", 9),
        )
        style.configure(
            "Accent.TButton",
            background=self.ACCENT,
            foreground="#001316",
            bordercolor=self.ACCENT,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT_DARK,
            relief="flat",
            padding=(18, 10),
            font=("Segoe UI Semibold", 10),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", self.BORDER), ("active", "#34D9E5")],
            foreground=[("disabled", self.MUTED), ("active", "#001316")],
        )
        style.configure(
            "Ghost.TButton",
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.PANEL_ALT,
            darkcolor=self.PANEL_ALT,
            relief="flat",
            padding=(12, 8),
            font=("Segoe UI", 9),
        )
        style.map(
            "Ghost.TButton",
            background=[("active", self.PANEL_HOVER), ("disabled", self.PANEL)],
            foreground=[("disabled", self.MUTED)],
        )
        style.configure(
            "Compact.TButton",
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            bordercolor=self.BORDER,
            lightcolor=self.PANEL_ALT,
            darkcolor=self.PANEL_ALT,
            relief="flat",
            padding=(8, 7),
            font=("Segoe UI", 8),
        )
        style.map(
            "Compact.TButton",
            background=[("active", self.PANEL_HOVER)],
        )
        style.configure(
            "Dark.TCombobox",
            fieldbackground=self.PANEL_ALT,
            background=self.PANEL_ALT,
            foreground=self.TEXT,
            arrowcolor=self.ACCENT,
            bordercolor=self.BORDER,
            lightcolor=self.BORDER,
            darkcolor=self.BORDER,
            padding=5,
        )
        style.map(
            "Dark.TCombobox",
            fieldbackground=[("readonly", self.PANEL_ALT)],
            foreground=[("readonly", self.TEXT)],
            selectbackground=[("readonly", self.PANEL_ALT)],
            selectforeground=[("readonly", self.TEXT)],
        )
        style.configure(
            "Cyber.Horizontal.TProgressbar",
            troughcolor=self.PANEL,
            background=self.ACCENT,
            bordercolor=self.BORDER,
            lightcolor=self.ACCENT,
            darkcolor=self.ACCENT_DARK,
            thickness=5,
        )

    # ------------------------------------------------------------------ layout
    def _build_ui(self):
        self._build_header()

        body = tk.Frame(self, bg=self.BG)
        body.pack(fill="both", expand=True, padx=24, pady=(18, 20))
        body.grid_columnconfigure(0, weight=6, minsize=520)
        body.grid_columnconfigure(1, weight=5, minsize=410)
        body.grid_rowconfigure(0, weight=1)

        config_panel = tk.Frame(body, bg=self.PANEL, highlightthickness=1, highlightbackground=self.BORDER)
        config_panel.grid(row=0, column=0, sticky="nsew", padx=(0, 12))
        config_panel.grid_columnconfigure(0, weight=1)

        console_panel = tk.Frame(body, bg=self.PANEL, highlightthickness=1, highlightbackground=self.BORDER)
        console_panel.grid(row=0, column=1, sticky="nsew", padx=(12, 0))
        console_panel.grid_rowconfigure(3, weight=1)
        console_panel.grid_columnconfigure(0, weight=1)

        self._build_configuration(config_panel)
        self._build_console(console_panel)

    def _build_header(self):
        header = tk.Canvas(self, height=112, bg=self.BG, highlightthickness=0)
        header.pack(fill="x")
        header.bind("<Configure>", lambda event: self._draw_header(header, event.width))

    def _draw_header(self, canvas: tk.Canvas, width: int):
        canvas.delete("all")
        canvas.create_polygon(
            0, 0,
            width, 0,
            width, 84,
            width - 38, 112,
            38, 112,
            0, 84,
            fill=self.PANEL_ALT,
            outline=self.BORDER,
        )
        canvas.create_line(42, 111, width - 42, 111, fill=self.ACCENT, width=2)
        canvas.create_polygon(0, 0, 210, 0, 180, 12, 0, 12, fill=self.ACCENT_DARK, outline="")
        canvas.create_text(
            44, 43,
            anchor="w",
            text="PROFILE // DOCS CHECKER",
            fill=self.TEXT,
            font=("Segoe UI Semibold", 19),
        )
        canvas.create_text(
            45, 73,
            anchor="w",
            text="DOCUMENT CONSISTENCY CONTROL",
            fill=self.ACCENT,
            font=("Consolas", 9, "bold"),
        )
        canvas.create_text(
            width - 45, 50,
            anchor="e",
            text="PDF  •  EXCEL  •  WORD",
            fill=self.MUTED,
            font=("Segoe UI", 10),
        )
        canvas.create_text(
            width - 45, 74,
            anchor="e",
            text="v1.3.0",
            fill=self.ACCENT,
            font=("Consolas", 9),
        )

    def _build_configuration(self, parent: tk.Frame):
        title_row = tk.Frame(parent, bg=self.PANEL)
        title_row.pack(fill="x", padx=20, pady=(18, 8))
        tk.Label(
            title_row,
            text="VALIDATION INPUTS",
            bg=self.PANEL,
            fg=self.TEXT,
            font=("Segoe UI Semibold", 12),
        ).pack(side="left")
        tk.Label(
            title_row,
            text="READY",
            bg=self.PANEL_ALT,
            fg=self.SUCCESS,
            padx=10,
            pady=4,
            font=("Consolas", 8, "bold"),
        ).pack(side="right")

        separator = tk.Frame(parent, height=1, bg=self.BORDER)
        separator.pack(fill="x", padx=20, pady=(0, 14))

        fields = tk.Frame(parent, bg=self.PANEL)
        fields.pack(fill="both", expand=True, padx=20)
        fields.grid_columnconfigure(0, weight=1)

        self._path_field(
            fields, 0,
            "EXCEL MASTER LIST  •  REQUIRED",
            self.excel_file_var,
            self.choose_excel_file,
            file_hint="Fixed German master template (.xlsx / .xlsm)",
        )

        sheet_row = tk.Frame(fields, bg=self.PANEL)
        sheet_row.grid(row=1, column=0, sticky="ew", pady=(0, 14))
        sheet_row.grid_columnconfigure(0, weight=1)
        tk.Label(sheet_row, text="EXCEL SHEET", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.sheet_combo = ttk.Combobox(sheet_row, textvariable=self.sheet_var, state="readonly", style="Dark.TCombobox")
        self.sheet_combo.grid(row=1, column=0, sticky="ew")
        ttk.Button(sheet_row, text="RELOAD", style="Compact.TButton", command=self.load_sheets).grid(row=1, column=1, padx=(8, 0))

        self._path_field(
            fields, 2,
            "WORD CHANGE NOTICE  •  OPTIONAL",
            self.word_file_var,
            self.choose_word_file,
            clear_command=lambda: self.word_file_var.set(""),
            file_hint="Selecting a Word file automatically enables Word-vs-Excel and page-count checks.",
        )

        self._path_field(
            fields, 3,
            "PDF DOCUMENT FOLDER  •  OPTIONAL",
            self.pdf_folder_var,
            self.choose_pdf_folder,
            clear_command=lambda: self.pdf_folder_var.set(""),
            file_hint="Required for PDF metadata checks and the true page count. The metadata table is read from page 2.",
        )

        issue_row = tk.Frame(fields, bg=self.PANEL)
        issue_row.grid(row=4, column=0, sticky="ew", pady=(0, 14))
        issue_row.grid_columnconfigure(0, weight=1)
        tk.Label(issue_row, text="EXPECTED AUSGABE", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        self.ausgabe_entry = self._dark_entry(issue_row, self.ausgabe_var)
        self.ausgabe_entry.grid(row=1, column=0, sticky="ew")
        tk.Label(issue_row, text="Month and year, e.g. 07.2026. Only required when PDFs are selected.", bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8)).grid(row=2, column=0, sticky="w", pady=(5, 0))

        self._path_field(
            fields, 5,
            "OUTPUT REPORT",
            self.output_file_var,
            self.choose_output_file,
            file_hint="Excel validation report (.xlsx)",
            browse_text="SAVE AS",
        )



        action_row = tk.Frame(parent, bg=self.PANEL)
        action_row.pack(fill="x", padx=20, pady=(0, 14), before=fields)
        self.run_button = ttk.Button(action_row, text="RUN VALIDATION", style="Accent.TButton", command=self.run)
        self.run_button.pack(side="left")
        self.open_button = ttk.Button(action_row, text="OPEN REPORT", style="Ghost.TButton", command=self.open_report, state="disabled")
        self.open_button.pack(side="left", padx=(9, 0))
        ttk.Button(action_row, text="CLEAR LOG", style="Ghost.TButton", command=self.clear_log).pack(side="right")

    def _build_console(self, parent: tk.Frame):
        top = tk.Frame(parent, bg=self.PANEL)
        top.grid(row=0, column=0, sticky="ew", padx=16, pady=(16, 8))
        tk.Label(top, text="SYSTEM LOG", bg=self.PANEL, fg=self.TEXT, font=("Segoe UI Semibold", 11)).pack(side="left")
        self.activity_label = tk.Label(top, text="IDLE", bg=self.PANEL_ALT, fg=self.MUTED, padx=9, pady=3, font=("Consolas", 8, "bold"))
        self.activity_label.pack(side="right")

        self.progress = ttk.Progressbar(parent, mode="determinate", value=0, maximum=100, style="Cyber.Horizontal.TProgressbar")
        self.progress.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 8))

        mode_strip = tk.Frame(parent, bg=self.PANEL_ALT, highlightthickness=1, highlightbackground=self.BORDER)
        mode_strip.grid(row=2, column=0, sticky="ew", padx=16, pady=(0, 10))
        tk.Label(mode_strip, text="MODE", bg=self.PANEL_ALT, fg=self.ACCENT, font=("Consolas", 8, "bold")).pack(side="left", padx=(10, 8), pady=7)
        tk.Label(mode_strip, textvariable=self.mode_text_var, bg=self.PANEL_ALT, fg=self.TEXT, font=("Segoe UI", 8), wraplength=420, justify="left").pack(side="left", fill="x", expand=True, padx=(0, 10), pady=7)

        text_frame = tk.Frame(parent, bg=self.BG, highlightthickness=1, highlightbackground=self.BORDER)
        text_frame.grid(row=3, column=0, sticky="nsew", padx=16, pady=(0, 16))
        text_frame.grid_rowconfigure(0, weight=1)
        text_frame.grid_columnconfigure(0, weight=1)

        self.output = tk.Text(
            text_frame,
            wrap="word",
            bg=self.BG,
            fg="#BFD0D4",
            insertbackground=self.ACCENT,
            selectbackground=self.ACCENT_DARK,
            selectforeground=self.TEXT,
            relief="flat",
            borderwidth=0,
            padx=12,
            pady=12,
            font=("Consolas", 9),
        )
        self.output.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.output.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.output.configure(yscrollcommand=scrollbar.set)
        self.output.tag_configure("accent", foreground=self.ACCENT)
        self.output.tag_configure("success", foreground=self.SUCCESS)
        self.output.tag_configure("warning", foreground=self.WARNING)
        self.output.tag_configure("error", foreground=self.ERROR)
        self.log("Profile Docs Checker initialized.", "accent")
        self.log("Choose Excel and at least one source: Word, PDFs, or both.")

    def _dark_entry(self, parent, variable: tk.StringVar) -> tk.Entry:
        return tk.Entry(
            parent,
            textvariable=variable,
            bg=self.PANEL_ALT,
            fg=self.TEXT,
            insertbackground=self.ACCENT,
            selectbackground=self.ACCENT_DARK,
            selectforeground=self.TEXT,
            relief="flat",
            highlightthickness=1,
            highlightbackground=self.BORDER,
            highlightcolor=self.ACCENT,
            disabledbackground=self.PANEL_ALT,
            disabledforeground=self.MUTED,
            font=("Segoe UI", 9),
        )

    def _path_field(
        self,
        parent,
        row: int,
        label: str,
        variable: tk.StringVar,
        browse_command,
        file_hint: str,
        clear_command=None,
        browse_text: str = "BROWSE",
    ):
        frame = tk.Frame(parent, bg=self.PANEL)
        frame.grid(row=row, column=0, sticky="ew", pady=(0, 14))
        frame.grid_columnconfigure(0, weight=1)
        tk.Label(frame, text=label, bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
        entry = self._dark_entry(frame, variable)
        entry.grid(row=1, column=0, sticky="ew")
        ttk.Button(frame, text=browse_text, style="Compact.TButton", command=browse_command).grid(row=1, column=1, padx=(8, 0))
        if clear_command:
            ttk.Button(frame, text="CLEAR", style="Compact.TButton", command=clear_command).grid(row=1, column=2, padx=(6, 0))
        tk.Label(frame, text=file_hint, bg=self.PANEL, fg=self.MUTED, font=("Segoe UI", 8), wraplength=430, justify="left").grid(row=2, column=0, columnspan=3, sticky="w", pady=(5, 0))

    # ------------------------------------------------------------------ file actions
    def choose_pdf_folder(self):
        path = filedialog.askdirectory(title="Choose folder containing PDFs")
        if path:
            self.pdf_folder_var.set(path)
            self._suggest_output(Path(path))

    def choose_excel_file(self):
        path = filedialog.askopenfilename(
            title="Choose Excel master list",
            filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")],
        )
        if path:
            self.excel_file_var.set(path)
            self.load_sheets()
            self._suggest_output(Path(path).parent)

    def choose_word_file(self):
        path = filedialog.askopenfilename(
            title="Choose Word change notice",
            filetypes=[("Word files", "*.docx"), ("All files", "*.*")],
        )
        if path:
            self.word_file_var.set(path)
            self._suggest_output(Path(path).parent)

    def choose_output_file(self):
        initial = self.output_file_var.get().strip() or "profile_docs_validation_report.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save validation report as",
            defaultextension=".xlsx",
            initialdir=str(Path(initial).parent) if Path(initial).parent.exists() else None,
            initialfile=Path(initial).name,
            filetypes=[("Excel report", "*.xlsx")],
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
            self.log("Excel sheets loaded: " + ", ".join(sheets), "success")
        except Exception as exc:
            messagebox.showerror("Could not load Excel sheets", str(exc))

    # ------------------------------------------------------------------ state
    def _refresh_mode_status(self):
        pdf = bool(self.pdf_folder_var.get().strip())
        word = bool(self.word_file_var.get().strip())
        if pdf and word:
            text = "FULL CONTROL  •  PDF ↔ Excel  •  Word ↔ Excel  •  Word page count ↔ actual PDF pages"
        elif pdf:
            text = "PDF CONTROL  •  PDF metadata ↔ Excel  •  Ausgabe validation"
        elif word:
            text = "WORD CONTROL  •  Word ↔ Excel  •  page-count result unavailable without PDFs"
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

    # ------------------------------------------------------------------ run
    def run(self):
        try:
            pdf_folder, excel, word_file, output, ausgabe = self._collect_args()
        except Exception as exc:
            messagebox.showerror("Missing or invalid input", str(exc))
            return

        if word_file and not pdf_folder:
            proceed = messagebox.askyesno(
                "PDF page count unavailable",
                "No PDF folder was selected.\n\n"
                "The tool will compare the Word table with Excel, but it cannot verify the true number of pages. "
                "This limitation will also be recorded in the report.\n\nContinue?",
                icon="warning",
            )
            if not proceed:
                return

        self.open_button.configure(state="disabled")
        self.run_button.configure(state="disabled")
        self.activity_label.configure(text="RUNNING", fg=self.ACCENT)
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)
        self.log("", None)
        self.log("Starting validation sequence...", "accent")

        args = (pdf_folder, excel, word_file, output, ausgabe)
        threading.Thread(target=self._run_thread, args=args, daemon=True).start()

    def _run_thread(self, pdf_folder, excel, word_file, output, ausgabe):
        try:
            def progress(message: str):
                self.after(0, lambda m=message: self.log(m))

            results = run_validation(
                pdf_folder=pdf_folder,
                excel_path=excel,
                output_path=output,
                expected_ausgabe=ausgabe,
                sheet_name=self.sheet_var.get().strip() or None,
                page_number=2,
                word_docx_path=word_file,
                progress_callback=progress,
            )
            counts = summarize_results(results)
            if counts:
                self.after(0, lambda: self.log("PDF result summary:", "accent"))
                for status, count in counts.items():
                    self.after(0, lambda s=status, c=count: self.log(f"  {s}: {c}"))
            else:
                self.after(0, lambda: self.log("PDF validation skipped; Word/Excel report created.", "warning"))
            self.after(0, lambda: self._finish_success(output))
        except Exception as exc:
            error_text = str(exc)
            self.after(0, lambda: self._finish_error(error_text))

    def _finish_success(self, output: str):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=100)
        self.run_button.configure(state="normal")
        self.open_button.configure(state="normal")
        self.activity_label.configure(text="COMPLETE", fg=self.SUCCESS)
        self.log(f"Report saved: {output}", "success")
        messagebox.showinfo("Validation complete", f"Report saved to:\n{output}")

    def _finish_error(self, error_text: str):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.run_button.configure(state="normal")
        self.activity_label.configure(text="ERROR", fg=self.ERROR)
        self.log("Validation failed: " + error_text, "error")
        messagebox.showerror("Validation failed", error_text)

    def log(self, text: str, tag: str | None = None):
        if tag:
            self.output.insert("end", text + "\n", tag)
        else:
            self.output.insert("end", text + "\n")
        self.output.see("end")
        self.update_idletasks()

    def clear_log(self):
        self.output.delete("1.0", "end")
        self.log("Log cleared.", "accent")

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
