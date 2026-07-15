from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import (
    PDF_FIELDS,
    PDF_FIELD_DISPLAY,
    preview_extraction,
    run_validation,
    summarize_results,
)


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Validate fixed-template PDF document-identification tables and/or a standard "
            "Word Engineering Change Notice against an Excel master list. When both Word "
            "and PDFs are supplied, Word Blatt/sheets is compared with the actual PDF page count."
        )
    )
    parser.add_argument("--pdf-folder", default=None, help="Optional folder containing PDF files.")
    parser.add_argument("--excel", required=True, help="Excel master list (.xlsx).")
    parser.add_argument("--out", required=True, help="Output validation report (.xlsx).")
    parser.add_argument(
        "--ausgabe",
        default=None,
        help="Expected Ausgabe as MM.YYYY. Required only when --pdf-folder is used; 07-2026 and 07/2026 are also accepted.",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet name. Default: first sheet.")
    parser.add_argument("--page", type=int, default=2, help="Visible PDF metadata page number. Default: 2.")
    parser.add_argument("--table-index", type=int, default=None, help="Optional 1-based PDF table number. Default: auto.")
    parser.add_argument("--row-start", type=int, default=None, help="Optional 1-based metadata row start. Default: auto.")
    parser.add_argument("--word", default=None, help="Optional standard Word Freigabe-/Änderungsmitteilung (.docx).")
    parser.add_argument("--word-mapping", default=None, help="Advanced: optional mapping JSON for a non-standard Word table.")
    parser.add_argument("--preview", action="store_true", help="Preview PDF extraction instead of running validation.")
    parser.add_argument("--preview-limit", type=int, default=5, help="Number of PDFs to preview. Default: 5.")
    args = parser.parse_args()

    if args.preview:
        if not args.pdf_folder:
            parser.error("--preview requires --pdf-folder")
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
            print(
                f"Total pages: {item.total_pages} | Metadata page: {item.page_number} | "
                f"Table: {item.table_index} | Row start: {item.row_start}"
            )
            for field in PDF_FIELDS:
                print(
                    f"{PDF_FIELD_DISPLAY[field]:24} raw={item.values_raw.get(field, '')!r}  "
                    f"normalized={item.values_norm.get(field, '')!r}"
                )
            if item.warnings:
                print("Warnings:", "; ".join(item.warnings))
        return 0

    def progress(message: str):
        print(message)

    results = run_validation(
        pdf_folder=Path(args.pdf_folder) if args.pdf_folder else None,
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
    if results:
        print("\nPDF Summary:")
        for status, count in summarize_results(results).items():
            print(f"  {status}: {count}")
    else:
        print("\nPDF validation skipped (no PDF folder supplied).")
    print(f"\nReport: {args.out}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
