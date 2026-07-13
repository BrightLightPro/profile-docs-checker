from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import preview_extraction, run_validation, summarize_results, PDF_FIELDS, PDF_FIELD_DISPLAY


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Validate fixed-template PDF document identification tables against an Excel master list. Optionally compare a Word table with the same Excel list."
    )
    parser.add_argument("--pdf-folder", required=True, help="Folder containing PDF files.")
    parser.add_argument("--excel", required=True, help="Excel master list (.xlsx).")
    parser.add_argument("--out", required=True, help="Output validation report (.xlsx).")
    parser.add_argument("--ausgabe", required=True, help="Expected Ausgabe as MM.YYYY, e.g. 07.2026. Also accepts 07-2026 and 07/2026.")
    parser.add_argument("--sheet", default=None, help="Excel sheet name. Default: first sheet.")
    parser.add_argument("--page", type=int, default=2, help="Visible PDF page number containing the table. Default: 2.")
    parser.add_argument("--table-index", type=int, default=None, help="Optional 1-based table number on the PDF page. Default: auto.")
    parser.add_argument("--row-start", type=int, default=None, help="Optional 1-based row in detected PDF table where Dokumentname starts. Default: auto.")
    parser.add_argument("--word", default=None, help="Optional Word .docx Freigabe-/Änderungsmitteilung. The standard template is recognized automatically.")
    parser.add_argument("--word-mapping", default=None, help="Optional custom mapping JSON for a different Word table template. Do not use it for the standard change-notice template.")
    parser.add_argument("--preview", action="store_true", help="Preview PDF extraction from the first few PDFs instead of running validation.")
    parser.add_argument("--preview-limit", type=int, default=5, help="Number of PDFs to preview. Default: 5.")
    args = parser.parse_args()

    if args.preview:
        previews = preview_extraction(
            args.pdf_folder,
            page_number=args.page,
            table_index=args.table_index,
            row_start=args.row_start,
            limit=args.preview_limit,
        )
        for item in previews:
            print("=" * 80)
            print(item.file_name)
            if item.error:
                print("ERROR:", item.error)
                continue
            print(f"Page: {item.page_number} | Table: {item.table_index} | Row start: {item.row_start}")
            for field in PDF_FIELDS:
                print(f"{PDF_FIELD_DISPLAY[field]:24} raw={item.values_raw.get(field, '')!r}  normalized={item.values_norm.get(field, '')!r}")
            if item.warnings:
                print("Warnings:", "; ".join(item.warnings))
        return 0

    def progress(msg: str):
        print(msg)

    results = run_validation(
        pdf_folder=Path(args.pdf_folder),
        excel_path=Path(args.excel),
        output_path=Path(args.out),
        expected_ausgabe=args.ausgabe,
        sheet_name=args.sheet,
        page_number=args.page,
        table_index=args.table_index,
        row_start=args.row_start,
        word_docx_path=Path(args.word) if args.word else None,
        word_mapping_path=Path(args.word_mapping) if args.word_mapping else None,
        progress_callback=progress,
    )
    print("\nPDF Summary:")
    for status, count in summarize_results(results).items():
        print(f"  {status}: {count}")
    print(f"\nReport: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
