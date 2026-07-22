from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import fill_word_copy_from_excel, run_validation, summarize_results


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Validate fixed-template PDF and Word metadata against an Excel master list, "
            "or fill the standard Word Engineering Change Notice from Excel."
        )
    )
    parser.add_argument("--pdf-folder", default=None, help="Optional folder containing PDF files.")
    parser.add_argument("--excel", required=True, help="Excel master list (.xlsx or .xlsm).")
    parser.add_argument("--out", required=True, help="Output report (.xlsx) or filled Word copy (.docx).")
    parser.add_argument(
        "--ausgabe", default=None,
        help="Expected Ausgabe as month/year. Required for validation with --pdf-folder.",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet name. Default: first sheet.")
    parser.add_argument("--word", default=None, help="Optional standard Word change notice (.docx).")
    parser.add_argument(
        "--fill-word", action="store_true",
        help="Fill the selected --word file from Excel instead of validating it.",
    )
    args = parser.parse_args()

    def progress(message: str):
        print(message)

    if args.fill_word:
        if not args.word:
            parser.error("--fill-word requires --word")
        result = fill_word_copy_from_excel(
            word_docx_path=Path(args.word), excel_path=Path(args.excel),
            output_path=Path(args.out), sheet_name=args.sheet,
            pdf_folder=Path(args.pdf_folder) if args.pdf_folder else None,
            progress_callback=progress,
        )
        print(f"\nFilled rows: {result.filled_rows}")
        print(f"Added rows: {result.added_rows}")
        print(f"Output: {result.output_path}")
        for warning in result.warnings:
            print(f"WARNING: {warning}")
        return 0

    results = run_validation(
        pdf_folder=Path(args.pdf_folder) if args.pdf_folder else None,
        excel_path=Path(args.excel), output_path=Path(args.out),
        expected_ausgabe=args.ausgabe, sheet_name=args.sheet, page_number=2,
        word_docx_path=Path(args.word) if args.word else None,
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
