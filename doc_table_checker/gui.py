from __future__ import annotations

import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from openpyxl import load_workbook

from .core import (
    PDF_FIELDS,
    PDF_FIELD_DISPLAY,
    preview_extraction,
    preview_word_extraction,
    run_validation,
    summarize_results,
    write_word_inspection_files,
)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Profile Docs Checker")
        self.geometry("1050x850")
        self.minsize(850, 600)

        self.pdf_folder_var = tk.StringVar()
        self.excel_file_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.output_file_var = tk.StringVar()
        self.ausgabe_var = tk.StringVar()
        self.page_var = tk.StringVar(value="2")
        self.table_index_var = tk.StringVar()
        self.row_start_var = tk.StringVar()
        self.use_word_var = tk.BooleanVar(value=False)
        self.word_file_var = tk.StringVar()
        self.word_mapping_var = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        root = ttk.Frame(self, padding=12)
        root.pack(fill="both", expand=True)

        title = ttk.Label(root, text="Profile Docs Checker", font=("Segoe UI", 16, "bold"))
        title.pack(anchor="w", pady=(0, 10))

        form = ttk.Frame(root)
        form.pack(fill="x")
        form.columnconfigure(1, weight=1)

        self._add_path_row(form, 0, "PDF folder", self.pdf_folder_var, self.choose_pdf_folder)
        self._add_path_row(form, 1, "Excel file", self.excel_file_var, self.choose_excel_file)

        ttk.Label(form, text="Excel sheet").grid(row=2, column=0, sticky="w", padx=(0, 8), pady=4)
        self.sheet_combo = ttk.Combobox(form, textvariable=self.sheet_var, values=[], state="normal")
        self.sheet_combo.grid(row=2, column=1, sticky="ew", pady=4)
        ttk.Button(form, text="Load sheets", command=self.load_sheets).grid(row=2, column=2, sticky="ew", padx=(8, 0), pady=4)

        ttk.Label(form, text="Expected Ausgabe (MM.YYYY)").grid(row=3, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.ausgabe_var).grid(row=3, column=1, sticky="ew", pady=4)
        ttk.Label(form, text="Example: 07.2026 / 07-2026").grid(row=3, column=2, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(form, text="PDF page number").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(form, textvariable=self.page_var, width=12).grid(row=4, column=1, sticky="w", pady=4)
        ttk.Label(form, text="Visible page number. Default: 2").grid(row=4, column=2, sticky="w", padx=(8, 0), pady=4)

        advanced = ttk.LabelFrame(root, text="Advanced extraction settings (normally leave blank)", padding=10)
        advanced.pack(fill="x", pady=(12, 8))
        advanced.columnconfigure(1, weight=1)
        ttk.Label(advanced, text="Table number on page").grid(row=0, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(advanced, textvariable=self.table_index_var, width=12).grid(row=0, column=1, sticky="w", pady=4)
        ttk.Label(advanced, text="Blank = auto-detect").grid(row=0, column=2, sticky="w", padx=(8, 0), pady=4)

        ttk.Label(advanced, text="Row where Dokumentname starts").grid(row=1, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(advanced, textvariable=self.row_start_var, width=12).grid(row=1, column=1, sticky="w", pady=4)
        ttk.Label(advanced, text="Blank = auto-detect. Use 1 if Dokumentname is first row.").grid(row=1, column=2, sticky="w", padx=(8, 0), pady=4)

        self._add_path_row(form, 5, "Output report", self.output_file_var, self.choose_output_file, save=True)

        word = ttk.LabelFrame(root, text="Optional Word table comparison", padding=10)
        word.pack(fill="x", pady=(8, 8))
        word.columnconfigure(1, weight=1)
        ttk.Checkbutton(word, text="Also compare Word table against Excel", variable=self.use_word_var).grid(row=0, column=0, columnspan=3, sticky="w", pady=4)
        self._add_path_row(word, 1, "Word .docx file", self.word_file_var, self.choose_word_file)
        self._add_path_row(word, 2, "Custom mapping JSON (optional)", self.word_mapping_var, self.choose_word_mapping)
        ttk.Label(word, text="The Freigabe-/Änderungsmitteilung template is recognized automatically. Use a mapping only for a different Word template.").grid(row=3, column=0, columnspan=2, sticky="w", pady=4)
        ttk.Button(word, text="Inspect custom table...", command=self.inspect_word_file).grid(row=3, column=2, sticky="ew", padx=(8, 0), pady=4)

        buttons = ttk.Frame(root)
        buttons.pack(fill="x", pady=(12, 8))
        ttk.Button(buttons, text="Preview PDF extraction", command=self.preview).pack(side="left")
        ttk.Button(buttons, text="Preview Word extraction", command=self.preview_word).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Run full validation", command=self.run).pack(side="left", padx=(8, 0))
        ttk.Button(buttons, text="Quit", command=self.destroy).pack(side="right")

        self.output = tk.Text(root, wrap="word", height=20)
        self.output.pack(fill="both", expand=True, pady=(4, 0))
        self.output.insert("end", "Workflow:\n")
        self.output.insert("end", "1. Select the PDF folder.\n")
        self.output.insert("end", "2. Select the Excel file and sheet.\n")
        self.output.insert("end", "3. Enter the global Ausgabe as MM.YYYY.\n")
        self.output.insert("end", "4. Click Preview PDF extraction before running the full validation.\n")
        self.output.insert("end", "5. Optional: choose the standard Word change-notice file; no mapping JSON is required.\n")

    def _add_path_row(self, parent, row, label, var, command, save=False):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 8), pady=4)
        ttk.Entry(parent, textvariable=var).grid(row=row, column=1, sticky="ew", pady=4)
        ttk.Button(parent, text="Choose..." if not save else "Save as...", command=command).grid(row=row, column=2, sticky="ew", padx=(8, 0), pady=4)

    def choose_pdf_folder(self):
        path = filedialog.askdirectory(title="Choose folder containing PDFs")
        if path:
            self.pdf_folder_var.set(path)
            if not self.output_file_var.get():
                self.output_file_var.set(str(Path(path) / "validation_report.xlsx"))

    def choose_excel_file(self):
        path = filedialog.askopenfilename(title="Choose Excel file", filetypes=[("Excel files", "*.xlsx *.xlsm"), ("All files", "*.*")])
        if path:
            self.excel_file_var.set(path)
            self.load_sheets()

    def choose_word_file(self):
        path = filedialog.askopenfilename(title="Choose Word file", filetypes=[("Word files", "*.docx"), ("All files", "*.*")])
        if path:
            self.word_file_var.set(path)

    def choose_word_mapping(self):
        path = filedialog.askopenfilename(title="Choose Word mapping JSON", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
        if path:
            self.word_mapping_var.set(path)

    def inspect_word_file(self):
        word_file = self.word_file_var.get().strip()
        if not word_file:
            messagebox.showerror("Word file missing", "Please choose a Word .docx file first.")
            return
        base = Path(word_file).with_suffix("")
        structure_default = str(base.parent / f"{base.name}_table_structure.json")
        mapping_default = str(base.parent / f"{base.name}_mapping_template.json")
        try:
            structure_out, mapping_out = write_word_inspection_files(
                word_file,
                structure_out=structure_default,
                mapping_out=mapping_default,
                include_samples=False,
            )
            self.log(f"Word structure saved: {structure_out}")
            self.log(f"Word mapping template saved: {mapping_out}")
            messagebox.showinfo(
                "Word inspection complete",
                "Created structure and custom mapping JSON files. The mapping is only needed for a non-standard Word table."
            )
        except Exception as exc:
            messagebox.showerror("Word inspection failed", str(exc))

    def choose_output_file(self):
        initial = self.output_file_var.get() or "validation_report.xlsx"
        path = filedialog.asksaveasfilename(
            title="Save report as",
            defaultextension=".xlsx",
            initialfile=Path(initial).name,
            filetypes=[("Excel report", "*.xlsx")],
        )
        if path:
            self.output_file_var.set(path)

    def load_sheets(self):
        excel = self.excel_file_var.get().strip()
        if not excel:
            return
        try:
            wb = load_workbook(excel, read_only=True, data_only=True)
            sheets = wb.sheetnames
            wb.close()
            self.sheet_combo["values"] = sheets
            if sheets and not self.sheet_var.get():
                self.sheet_var.set(sheets[0])
            self.log(f"Loaded sheets: {', '.join(sheets)}")
        except Exception as exc:
            messagebox.showerror("Could not load Excel sheets", str(exc))

    def parse_int_or_none(self, value: str, name: str):
        value = value.strip()
        if not value:
            return None
        try:
            v = int(value)
        except ValueError:
            raise ValueError(f"{name} must be a number or blank.")
        if v < 1:
            raise ValueError(f"{name} must be 1 or higher.")
        return v

    def get_common_args(self):
        pdf_folder = self.pdf_folder_var.get().strip()
        excel = self.excel_file_var.get().strip()
        output = self.output_file_var.get().strip()
        ausgabe = self.ausgabe_var.get().strip()
        if not pdf_folder:
            raise ValueError("Please choose a PDF folder.")
        if not excel:
            raise ValueError("Please choose an Excel file.")
        if not ausgabe:
            raise ValueError("Please enter Expected Ausgabe as MM.YYYY, for example 07.2026.")
        try:
            page = int(self.page_var.get().strip() or "2")
        except ValueError:
            raise ValueError("PDF page number must be a number.")
        if page < 1:
            raise ValueError("PDF page number must be 1 or higher.")
        table_index = self.parse_int_or_none(self.table_index_var.get(), "Table number")
        row_start = self.parse_int_or_none(self.row_start_var.get(), "Row start")
        word_file = self.word_file_var.get().strip() if self.use_word_var.get() else None
        word_mapping = self.word_mapping_var.get().strip() if self.use_word_var.get() else None
        if self.use_word_var.get() and not word_file:
            raise ValueError("Word comparison is enabled, but no Word .docx file was selected.")
        return pdf_folder, excel, output, ausgabe, page, table_index, row_start, word_file, word_mapping

    def log(self, text: str):
        self.output.insert("end", text + "\n")
        self.output.see("end")
        self.update_idletasks()

    def preview(self):
        try:
            pdf_folder = self.pdf_folder_var.get().strip()
            if not pdf_folder:
                raise ValueError("Please choose a PDF folder.")
            try:
                page = int(self.page_var.get().strip() or "2")
            except ValueError:
                raise ValueError("PDF page number must be a number.")
            if page < 1:
                raise ValueError("PDF page number must be 1 or higher.")
            table_index = self.parse_int_or_none(self.table_index_var.get(), "Table number")
            row_start = self.parse_int_or_none(self.row_start_var.get(), "Row start")

            self.log("\nPreview extraction...")
            previews = preview_extraction(pdf_folder, page_number=page, table_index=table_index, row_start=row_start, limit=5)
            for item in previews:
                self.log("-" * 70)
                self.log(item.file_name)
                if item.error:
                    self.log("ERROR: " + item.error)
                    continue
                self.log(f"Page {item.page_number}, table {item.table_index}, row start {item.row_start}")
                for field in PDF_FIELDS:
                    self.log(f"  {PDF_FIELD_DISPLAY[field]:22} raw={item.values_raw.get(field, '')!r}  norm={item.values_norm.get(field, '')!r}")
                if item.warnings:
                    self.log("  Warnings: " + "; ".join(item.warnings))
        except Exception as exc:
            messagebox.showerror("Preview failed", str(exc))

    def preview_word(self):
        try:
            word_file = self.word_file_var.get().strip()
            if not word_file:
                raise ValueError("Please choose a Word .docx file first.")
            mapping = self.word_mapping_var.get().strip() or None
            records, info = preview_word_extraction(word_file, mapping_path=mapping, limit=10)
            self.log("\nPreview Word extraction...")
            self.log(f"Template: {info.template_type}")
            self.log(f"Table index: {info.table_index_zero_based}; data starts at Word row {info.data_start_row_zero_based + 1}")
            self.log(f"Columns: {info.columns_by_index_zero_based}")
            self.log(f"Top-right Number: {info.change_number_raw!r}")
            if info.warnings:
                self.log("Warnings: " + "; ".join(info.warnings))
            for rec in records:
                self.log("-" * 70)
                self.log(f"Word row {rec.word_row_number}")
                for field in ("dokumentnummer", "artikelnummer", "revision", "version", "freigabe"):
                    if field in rec.values_raw:
                        self.log(f"  {PDF_FIELD_DISPLAY[field]:22} raw={rec.values_raw.get(field, '')!r}  norm={rec.values_norm.get(field, '')!r}")
            if not records:
                self.log("No data rows found.")
        except Exception as exc:
            messagebox.showerror("Word preview failed", str(exc))

    def run(self):
        try:
            args = self.get_common_args()
            if not args[2]:
                raise ValueError("Please choose an output report path.")
        except Exception as exc:
            messagebox.showerror("Missing or invalid input", str(exc))
            return

        self.log("\nStarting validation...")
        thread = threading.Thread(target=self._run_thread, args=args, daemon=True)
        thread.start()

    def _run_thread(self, pdf_folder, excel, output, ausgabe, page, table_index, row_start, word_file, word_mapping):
        try:
            def progress(msg: str):
                self.after(0, lambda m=msg: self.log(m))

            results = run_validation(
                pdf_folder=pdf_folder,
                excel_path=excel,
                output_path=output,
                expected_ausgabe=ausgabe,
                sheet_name=self.sheet_var.get().strip() or None,
                page_number=page,
                table_index=table_index,
                row_start=row_start,
                word_docx_path=word_file,
                word_mapping_path=word_mapping,
                progress_callback=progress,
            )
            counts = summarize_results(results)
            self.after(0, lambda: self.log("\nSummary:"))
            for status, count in counts.items():
                self.after(0, lambda s=status, c=count: self.log(f"  {s}: {c}"))
            self.after(0, lambda: messagebox.showinfo("Validation complete", f"Report saved to:\n{output}"))
        except Exception as exc:
            # Capture exception text now. Python clears the exception variable after
            # the except block, so a delayed Tkinter lambda must not reference `exc`
            # directly; otherwise the GUI can appear to do nothing after
            # "Starting validation...".
            error_text = str(exc)
            self.after(0, lambda e=error_text: self.log("ERROR: " + e))
            self.after(0, lambda e=error_text: messagebox.showerror("Validation failed", e))


def main():
    app = App()
    app.mainloop()
