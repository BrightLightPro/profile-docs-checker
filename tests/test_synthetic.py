from __future__ import annotations

from pathlib import Path
import shutil
import sys

# Allow running from repo root without installing package.
ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import fitz
from openpyxl import Workbook, load_workbook

from doc_table_checker.core import preview_extraction, run_validation, summarize_results


def make_table_pdf(path: Path, rows: list[tuple[str, str]]):
    doc = fitz.open()
    doc.new_page(width=595, height=842)  # page 1, irrelevant
    page = doc.new_page(width=595, height=842)  # page 2

    x0, y0 = 70, 120
    col1_w, col2_w = 190, 250
    row_h = 28
    rows_count = len(rows)
    x1 = x0 + col1_w + col2_w
    y1 = y0 + rows_count * row_h

    # Outer border and grid.
    page.draw_rect(fitz.Rect(x0, y0, x1, y1), width=0.8)
    page.draw_line((x0 + col1_w, y0), (x0 + col1_w, y1), width=0.8)
    for i in range(1, rows_count):
        y = y0 + i * row_h
        page.draw_line((x0, y), (x1, y), width=0.8)

    for i, (label, value) in enumerate(rows):
        y = y0 + i * row_h + 8
        page.insert_text((x0 + 8, y), label, fontsize=9)
        page.insert_text((x0 + col1_w + 8, y), value, fontsize=9)

    doc.save(path)
    doc.close()


def make_excel(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Master"
    # Put headers on row 3 to test header auto-detection.
    ws.append(["irrelevant"])
    ws.append(["also irrelevant"])
    ws.append(["Sprache", "Dok.-Nr.", "DOK-ID", "Freigabe-/ Änd.-Nr.", "Artikel-Nr.", "Rev.", "Vers.", "Ignored"])
    ws.append(["de", "DN-100", "DOC-001", "FR-1", "ART-123", "1", "2", "x"])
    ws.append(["en", "DN-200", "DOC-002", "FR-2", "ART-222", "1", "1", "x"])
    ws.append(["de", "DN-300", "DOC-DUP", "FR-3", "ART-333", "1", "1", "x"])
    ws.append(["de", "DN-301", "DOC-DUP", "FR-4", "ART-444", "2", "1", "x"])
    ws.append(["fr", "DN-400", "DOC-004", "FR-5", "ART-555", "3", "2", "x"])
    wb.save(path)


def make_dataset(base: Path):
    if base.exists():
        shutil.rmtree(base)
    pdf_dir = base / "pdfs"
    pdf_dir.mkdir(parents=True)
    excel = base / "master.xlsx"
    make_excel(excel)

    labels_de = [
        "Dokumentname", "Sprache", "Dokumentnummer", "DOK-ID", "Freigabe-/Änd.-Nr.",
        "Artikelnummer", "Revision", "Version", "Ausgabe"
    ]
    labels_en = [
        "Document name", "Language", "Document number", "DOC-ID", "Release/change no.",
        "Article number", "Revision", "Version", "Issue"
    ]

    def rows(labels, values):
        return list(zip(labels, values))

    make_table_pdf(pdf_dir / "01_ok_normalized.pdf", rows(labels_de, [
        "Ignored name", "de_DE", "DN-100", "DOC-001", "FR-1", "ART-123", "01", "02", "07.2026"
    ]))
    make_table_pdf(pdf_dir / "02_mismatch.pdf", rows(labels_en, [
        "Ignored name", "en_US", "DN-200", "DOC-002", "FR-2", "WRONG-ARTICLE", "1", "1", "07/2026"
    ]))
    make_table_pdf(pdf_dir / "03_likely_wrong_dok_id.pdf", rows(labels_en, [
        "Ignored name", "fr_FR", "DN-400", "DOC-DOES-NOT-EXIST", "FR-5", "ART-555", "03", "02", "2026-07"
    ]))
    make_table_pdf(pdf_dir / "04_duplicate_resolved.pdf", rows(labels_de, [
        "Ignored name", "de_DE", "DN-300", "DOC-DUP", "FR-3", "ART-333", "1", "1", "Juli 2026"
    ]))
    make_table_pdf(pdf_dir / "05_ausgabe_mismatch.pdf", rows(labels_de, [
        "Ignored name", "de_DE", "DN-100", "DOC-001", "FR-1", "ART-123", "1", "2", "08.2026"
    ]))
    return pdf_dir, excel


def run_synthetic_test():
    base = ROOT / "test_output"
    pdf_dir, excel = make_dataset(base)
    out = base / "validation_report.xlsx"

    previews = preview_extraction(pdf_dir, page_number=2, limit=2)
    assert previews[0].values_norm["sprache"] == "de", previews[0].values_norm
    assert previews[0].values_norm["ausgabe"] == "2026-07", previews[0].values_norm

    results = run_validation(pdf_dir, excel, out, expected_ausgabe="07.2026", sheet_name="Master", page_number=2)
    counts = summarize_results(results)
    print("Result counts:", counts)
    assert out.exists(), "Report was not created"
    assert counts.get("OK_WITH_NORMALIZATION", 0) >= 1, counts
    assert counts.get("MISMATCH", 0) == 1, counts
    assert counts.get("LIKELY_WRONG_DOK_ID", 0) == 1, counts
    assert counts.get("DUPLICATE_DOK_ID_RESOLVED", 0) == 1, counts
    assert counts.get("AUSGABE_MISMATCH", 0) == 1, counts

    wb = load_workbook(out, read_only=True, data_only=True)
    assert "Summary" in wb.sheetnames
    assert "Detailed comparison" in wb.sheetnames
    ws = wb["Summary"]
    assert ws.max_row >= 6
    wb.close()
    print("Synthetic test passed.")
    print("Sample report:", out)


if __name__ == "__main__":
    run_synthetic_test()
