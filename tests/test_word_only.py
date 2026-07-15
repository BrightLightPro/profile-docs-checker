from __future__ import annotations

from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(ROOT / "tests"))

from openpyxl import load_workbook

from doc_table_checker.core import run_validation
from test_synthetic import make_dataset
from test_change_notice_auto import make_change_notice


def run_word_only_test():
    base = ROOT / "test_output_word_only"
    _pdf_dir, excel = make_dataset(base)
    word = base / "change_notice_word_only.docx"
    make_change_notice(word)
    out = base / "word_only_report.xlsx"

    results = run_validation(
        pdf_folder=None,
        excel_path=excel,
        output_path=out,
        expected_ausgabe=None,
        sheet_name="Master",
        word_docx_path=word,
    )
    assert results == []
    assert out.exists()

    wb = load_workbook(out, read_only=True, data_only=True)
    assert wb.sheetnames == ["Overview", "Issues"]
    overview = list(wb["Overview"].iter_rows(min_row=10, values_only=True))
    assert any(row[14] == "Unavailable — no PDFs" for row in overview if row[6])
    issues = list(wb["Issues"].iter_rows(min_row=5, values_only=True))
    assert any(row[1] == "Run limitation" and row[5] == "Page-count verification" for row in issues)
    wb.close()
    print("Word-only validation test passed.")


if __name__ == "__main__":
    run_word_only_test()
