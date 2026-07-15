from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import run_validation, summarize_results


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Validate fixed-template PDF document-identification tables and/or the standard "
            "Word Engineering Change Notice against an Excel master list. When both Word "
            "and PDFs are supplied, Word Blatt/sheets is compared with the actual PDF page count."
        )
    )
    parser.add_argument("--pdf-folder", default=None, help="Optional folder containing PDF files.")
    parser.add_argument("--excel", required=True, help="Excel master list (.xlsx or .xlsm).")
    parser.add_argument("--out", required=True, help="Output validation report (.xlsx).")
    parser.add_argument(
        "--ausgabe",
        default=None,
        help="Expected Ausgabe as month/year. Required with --pdf-folder; e.g. 07.2026, 07-2026 or 07/2026.",
    )
    parser.add_argument("--sheet", default=None, help="Excel sheet name. Default: first sheet.")
    parser.add_argument("--word", default=None, help="Optional standard Word Freigabe-/Änderungsmitteilung (.docx).")
    args = parser.parse_args()

    def progress(message: str):
        print(message)

    results = run_validation(
        pdf_folder=Path(args.pdf_folder) if args.pdf_folder else None,
        excel_path=Path(args.excel),
        output_path=Path(args.out),
        expected_ausgabe=args.ausgabe,
        sheet_name=args.sheet,
        page_number=2,
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
