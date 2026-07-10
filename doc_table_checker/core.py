"""Core validation logic for the PDF/Excel document table checker.

The checker intentionally does NOT depend on PDF label text. It extracts values by
row position from the fixed table template on a configured PDF page.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, date
from pathlib import Path
import re
import unicodedata
import contextlib
import io
from typing import Dict, Iterable, List, Optional, Sequence, Tuple

import fitz  # PyMuPDF
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# Internal field names used throughout the program.
PDF_FIELDS = [
    "sprache",
    "dokumentnummer",
    "dok_id",
    "freigabe",
    "artikelnummer",
    "revision",
    "version",
    "ausgabe",
]

PDF_FIELD_DISPLAY = {
    "sprache": "Sprache",
    "dokumentnummer": "Dokumentnummer",
    "dok_id": "DOK-ID",
    "freigabe": "Freigabe-/Änd.-Nr.",
    "artikelnummer": "Artikelnummer",
    "revision": "Revision",
    "version": "Version",
    "ausgabe": "Ausgabe",
}

# The row order inside the fixed PDF table. Row 0 = Dokumentname, ignored.
PDF_ROW_INDEX = {
    "sprache": 1,
    "dokumentnummer": 2,
    "dok_id": 3,
    "freigabe": 4,
    "artikelnummer": 5,
    "revision": 6,
    "version": 7,
    "ausgabe": 8,
}

EXCEL_FIELD_TO_HEADER = {
    "sprache": "Sprache",
    "dokumentnummer": "Dok.-Nr.",
    "dok_id": "DOK-ID",
    "freigabe": "Freigabe-/ Änd.-Nr.",
    "artikelnummer": "Artikel-Nr.",
    "revision": "Rev.",
    "version": "Vers.",
}

MATCH_WEIGHTS = {
    "dokumentnummer": 40,
    "artikelnummer": 30,
    "freigabe": 20,
    "revision": 10,
    "version": 10,
    "sprache": 10,
}

MONTH_NAMES = {
    # German
    "januar": 1, "jan": 1,
    "februar": 2, "feb": 2,
    "maerz": 3, "märz": 3, "mar": 3, "mrz": 3,
    "april": 4, "apr": 4,
    "mai": 5,
    "juni": 6, "jun": 6,
    "juli": 7, "jul": 7,
    "august": 8, "aug": 8,
    "september": 9, "sep": 9, "sept": 9,
    "oktober": 10, "okt": 10, "oct": 10,
    "november": 11, "nov": 11,
    "dezember": 12, "dez": 12, "dec": 12,
    # English
    "january": 1,
    "february": 2,
    "march": 3,
    "may": 5,
    "june": 6,
    "july": 7,
    "october": 10,
    "december": 12,
    # French
    "janvier": 1, "fevrier": 2, "février": 2, "mars": 3, "avril": 4,
    "juin": 6, "juillet": 7, "aout": 8, "août": 8,
    "septembre": 9, "octobre": 10, "novembre": 11, "decembre": 12, "décembre": 12,
    # Spanish / Portuguese / Italian common month names
    "enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5,
    "junio": 6, "julio": 7, "agosto": 8, "septiembre": 9, "setiembre": 9,
    "octubre": 10, "noviembre": 11, "diciembre": 12,
    "janeiro": 1, "fevereiro": 2, "marco": 3, "março": 3, "maio": 5,
    "julho": 7, "setembro": 9, "outubro": 10, "dezembro": 12,
    "gennaio": 1, "febbraio": 2, "maggio": 5, "giugno": 6,
    "luglio": 7, "ottobre": 10, "dicembre": 12,
}

STATUS_ORDER = {
    "OK": 1,
    "OK_WITH_NORMALIZATION": 2,
    "AUSGABE_MISMATCH": 3,
    "MISMATCH": 4,
    "MISMATCH_AND_AUSGABE_MISMATCH": 5,
    "LIKELY_WRONG_DOK_ID": 6,
    "DUPLICATE_DOK_ID_RESOLVED": 7,
    "DUPLICATE_DOK_ID_AMBIGUOUS": 8,
    "DOK_ID_NOT_FOUND": 9,
    "NO_RELIABLE_MATCH": 10,
    "EXTRACTION_ERROR": 11,
}


@dataclass
class ExtractedPDF:
    file_path: Path
    file_name: str
    values_raw: Dict[str, str] = field(default_factory=dict)
    values_norm: Dict[str, str] = field(default_factory=dict)
    page_number: int = 2  # Human visible page number, 1-based.
    table_index: Optional[int] = None  # 1-based within detected tables.
    row_start: Optional[int] = None  # 1-based row where Dokumentname starts within detected table.
    error: Optional[str] = None
    warnings: List[str] = field(default_factory=list)


@dataclass
class ExcelRecord:
    row_number: int
    values_raw: Dict[str, str]
    values_norm: Dict[str, str]


@dataclass
class FieldComparison:
    file_name: str
    field: str
    source: str  # "Excel" or "User input"
    pdf_raw: str
    pdf_norm: str
    reference_raw: str
    reference_norm: str
    result: str
    excel_row: Optional[int] = None


@dataclass
class ValidationResult:
    extracted: ExtractedPDF
    status: str
    confidence: float = 0.0
    matched_excel_row: Optional[int] = None
    issues: List[str] = field(default_factory=list)
    comparisons: List[FieldComparison] = field(default_factory=list)
    candidate_rows: List[Tuple[int, float]] = field(default_factory=list)


def clean_text(value: object) -> str:
    """Return a stable one-line string without changing semantic content too much."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, date):
        return value.strftime("%d.%m.%Y")
    text = str(value)
    text = unicodedata.normalize("NFKC", text)
    text = text.replace("\u00a0", " ")
    text = text.replace("\r", " ").replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_identifier(value: object) -> str:
    text = clean_text(value)
    # Normalize dash variants and spaces around common separators.
    text = text.replace("–", "-").replace("—", "-").replace("−", "-")
    text = re.sub(r"\s*([/\\\-_:])\s*", r"\1", text)
    text = re.sub(r"\s+", " ", text).strip()
    # Case-insensitive matching for IDs/codes while keeping reported raw values unchanged.
    return text.casefold()


def normalize_header(value: object) -> str:
    text = clean_text(value).casefold()
    text = text.replace("ä", "ae").replace("ö", "oe").replace("ü", "ue").replace("ß", "ss")
    return re.sub(r"[^a-z0-9]+", "", text)


def normalize_language(value: object) -> str:
    text = clean_text(value).casefold().replace("-", "_")
    text_no_space = re.sub(r"\s+", "", text)
    # PDF: de_DE, en_US, pt_BR. Excel: de, en, pt.
    m = re.match(r"^([a-z]{2})(?:_[a-z]{2})?$", text_no_space)
    if m:
        return m.group(1)
    if re.match(r"^[a-z]{2}$", text_no_space):
        return text_no_space
    # Some PDF extractors split underscores oddly, e.g. "de DE _".
    pairs = re.findall(r"[a-z]{2}", text)
    if pairs:
        return pairs[0]
    return text_no_space


def normalize_revision_version(value: object) -> str:
    text = normalize_identifier(value)
    # Common safe normalization: 01 == 1, 002 == 2 for pure numeric revisions/versions.
    if re.fullmatch(r"0*\d+", text):
        return str(int(text))
    return text


def _normalize_month_number(month: int, year: int) -> Optional[str]:
    if 1 <= month <= 12 and 1900 <= year <= 2100:
        return f"{year:04d}-{month:02d}"
    return None


def normalize_month_year(value: object, allow_two_digit_year: bool = True) -> str:
    """Normalize many month/year formats to YYYY-MM, or return empty string if unsupported."""
    text = clean_text(value).casefold()
    if not text:
        return ""
    text = text.replace("–", "-").replace("—", "-").replace("−", "-")
    text = re.sub(r"\s+", " ", text).strip()

    # Date objects from Excel/openpyxl etc.
    if isinstance(value, (datetime, date)):
        return f"{value.year:04d}-{value.month:02d}"

    # Month-first, preferred: MM.YYYY / MM/YYYY / MM-YYYY / MM YYYY
    m = re.fullmatch(r"(\d{1,2})[.\-/\s](\d{4})", text)
    if m:
        return _normalize_month_number(int(m.group(1)), int(m.group(2))) or ""

    # Year-first fallback: YYYY-MM / YYYY/MM / YYYY.MM / YYYY MM
    m = re.fullmatch(r"(\d{4})[.\-/\s](\d{1,2})", text)
    if m:
        return _normalize_month_number(int(m.group(2)), int(m.group(1))) or ""

    if allow_two_digit_year:
        # Two-digit years are treated as 2000-2099. This is conservative for current technical docs.
        m = re.fullmatch(r"(\d{1,2})[.\-/\s](\d{2})", text)
        if m:
            return _normalize_month_number(int(m.group(1)), 2000 + int(m.group(2))) or ""

    # Month name + year, or year + month name.
    m = re.fullmatch(r"([a-zäöüßéèêûùàáíóúçãõ]+)\.?\s+(\d{4})", text)
    if m:
        month = MONTH_NAMES.get(m.group(1))
        if month:
            return _normalize_month_number(month, int(m.group(2))) or ""
    m = re.fullmatch(r"(\d{4})\s+([a-zäöüßéèêûùàáíóúçãõ]+)\.?,?", text)
    if m:
        month = MONTH_NAMES.get(m.group(2))
        if month:
            return _normalize_month_number(month, int(m.group(1))) or ""

    return ""


def normalize_expected_ausgabe(user_value: str) -> str:
    """Strict expected input: MM.YYYY or M.YYYY, normalized to YYYY-MM."""
    raw = clean_text(user_value)
    m = re.fullmatch(r"(\d{1,2})\.(\d{4})", raw)
    if not m:
        raise ValueError("Expected Ausgabe must be entered as MM.YYYY, for example 07.2026")
    normalized = _normalize_month_number(int(m.group(1)), int(m.group(2)))
    if not normalized:
        raise ValueError("Expected Ausgabe has an invalid month or year. Use MM.YYYY, for example 07.2026")
    return normalized


def normalize_field(field_name: str, value: object) -> str:
    if field_name == "sprache":
        return normalize_language(value)
    if field_name in {"revision", "version"}:
        return normalize_revision_version(value)
    if field_name == "ausgabe":
        return normalize_month_year(value)
    return normalize_identifier(value)


def excel_cell_to_text(cell) -> str:
    """Read Excel cell text while preserving obvious zero-padded numeric formats."""
    value = cell.value
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime("%d.%m.%Y")
    if isinstance(value, (int, float)):
        if isinstance(value, float) and value.is_integer():
            value = int(value)
        text = str(value)
        # Preserve simple zero-padded number formats like 000000 or 0000000.
        number_format = str(cell.number_format or "")
        if re.fullmatch(r"0+", number_format) and isinstance(value, int):
            text = str(value).zfill(len(number_format))
        return text
    return clean_text(value)


def load_excel_records(excel_path: Path | str, sheet_name: Optional[str] = None) -> Tuple[List[ExcelRecord], str, int, Dict[str, int]]:
    """Load relevant Excel columns from the fixed German template.

    Returns: records, actual_sheet_name, header_row_number, columns_by_field.
    """
    excel_path = Path(excel_path)
    if not excel_path.exists():
        raise FileNotFoundError(f"Excel file not found: {excel_path}")

    wb = load_workbook(excel_path, data_only=True, read_only=False)
    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(f"Sheet '{sheet_name}' not found. Available sheets: {', '.join(wb.sheetnames)}")
        ws = wb[sheet_name]
    else:
        ws = wb[wb.sheetnames[0]]

    header_variants = {normalize_header(header): field for field, header in EXCEL_FIELD_TO_HEADER.items()}
    # A few tolerant variants for the fixed template.
    header_variants.update({
        normalize_header("Dok Nr"): "dokumentnummer",
        normalize_header("Dokumentnummer"): "dokumentnummer",
        normalize_header("Dokument-Nr."): "dokumentnummer",
        normalize_header("Dok-ID"): "dok_id",
        normalize_header("DOK ID"): "dok_id",
        normalize_header("Freigabe-/Änd.-Nr."): "freigabe",
        normalize_header("Freigabe / Änd. Nr."): "freigabe",
        normalize_header("Artikel Nr."): "artikelnummer",
        normalize_header("Artikelnummer"): "artikelnummer",
        normalize_header("Revision"): "revision",
        normalize_header("Version"): "version",
    })

    best_header_row = None
    best_columns: Dict[str, int] = {}
    best_count = 0
    max_scan_row = min(ws.max_row, 50)
    for row in range(1, max_scan_row + 1):
        columns: Dict[str, int] = {}
        for col in range(1, ws.max_column + 1):
            key = normalize_header(ws.cell(row, col).value)
            if key in header_variants:
                columns[header_variants[key]] = col
        if len(columns) > best_count:
            best_count = len(columns)
            best_header_row = row
            best_columns = columns
        if best_count == len(EXCEL_FIELD_TO_HEADER):
            break

    missing = [EXCEL_FIELD_TO_HEADER[f] for f in EXCEL_FIELD_TO_HEADER if f not in best_columns]
    if best_header_row is None or missing:
        raise ValueError(
            "Could not find all required Excel columns. Missing: "
            + ", ".join(missing)
            + ". Check the sheet/template."
        )

    records: List[ExcelRecord] = []
    for row in range(best_header_row + 1, ws.max_row + 1):
        raw: Dict[str, str] = {}
        norm: Dict[str, str] = {}
        for field_name, col in best_columns.items():
            raw_value = excel_cell_to_text(ws.cell(row, col))
            raw[field_name] = raw_value
            norm[field_name] = normalize_field(field_name, raw_value)
        if any(raw.get(f, "") for f in EXCEL_FIELD_TO_HEADER):
            records.append(ExcelRecord(row_number=row, values_raw=raw, values_norm=norm))

    return records, ws.title, best_header_row, best_columns


def _last_nonempty_cell_after_label(row: Sequence[object]) -> str:
    """For a table row, take the last non-empty cell after the label column."""
    cells = [clean_text(c) for c in row]
    if len(cells) <= 1:
        return cells[0] if cells else ""
    right_side = [c for c in cells[1:] if c]
    if right_side:
        return right_side[-1]
    # If table extraction merged the row unexpectedly, return last non-empty cell.
    non_empty = [c for c in cells if c]
    return non_empty[-1] if non_empty else ""


def _candidate_score_from_rows(rows: Sequence[Sequence[object]], start: int) -> float:
    """Score a possible 9-row metadata table slice. Higher is better."""
    if start < 0 or start + 8 >= len(rows):
        return -1
    raw_values = {field_name: _last_nonempty_cell_after_label(rows[start + idx]) for field_name, idx in PDF_ROW_INDEX.items()}
    score = 0.0
    # Positive signals without relying on labels.
    if re.fullmatch(r"[A-Za-z]{2}[_-][A-Za-z]{2}", clean_text(raw_values.get("sprache", ""))):
        score += 35
    elif raw_values.get("sprache"):
        score += 10
    if raw_values.get("dok_id"):
        score += 25
    if raw_values.get("dokumentnummer"):
        score += 15
    if raw_values.get("artikelnummer"):
        score += 15
    if normalize_month_year(raw_values.get("ausgabe", "")):
        score += 15
    # Prefer slices with values in most relevant rows.
    score += sum(2 for v in raw_values.values() if clean_text(v))
    return score


def _extract_from_table_rows(
    rows: Sequence[Sequence[object]],
    file_path: Path,
    page_number: int,
    table_index: int,
    forced_row_start: Optional[int] = None,
) -> ExtractedPDF:
    if forced_row_start is not None:
        start = forced_row_start - 1
    else:
        possible = [(i, _candidate_score_from_rows(rows, i)) for i in range(0, max(0, len(rows) - 8))]
        if not possible:
            raise ValueError("Detected table has fewer than 9 rows.")
        start, _score = max(possible, key=lambda x: x[1])
    if start < 0 or start + 8 >= len(rows):
        raise ValueError(f"Row start {start + 1} is invalid for a table with {len(rows)} rows.")

    raw: Dict[str, str] = {}
    norm: Dict[str, str] = {}
    for field_name, idx in PDF_ROW_INDEX.items():
        value = _last_nonempty_cell_after_label(rows[start + idx])
        raw[field_name] = value
        norm[field_name] = normalize_field(field_name, value)

    warnings: List[str] = []
    if not norm.get("sprache"):
        warnings.append("Could not normalize Sprache from PDF.")
    if not raw.get("dok_id"):
        warnings.append("DOK-ID is empty in PDF extraction.")
    if raw.get("ausgabe") and not norm.get("ausgabe"):
        warnings.append("Could not normalize Ausgabe from PDF.")

    return ExtractedPDF(
        file_path=file_path,
        file_name=file_path.name,
        values_raw=raw,
        values_norm=norm,
        page_number=page_number,
        table_index=table_index,
        row_start=start + 1,
        warnings=warnings,
    )


def extract_pdf_table(
    pdf_path: Path | str,
    page_number: int = 2,
    table_index: Optional[int] = None,
    row_start: Optional[int] = None,
) -> ExtractedPDF:
    """Extract values from the fixed PDF table on a given visible page number.

    table_index and row_start are 1-based optional overrides. Leave them blank for auto detection.
    """
    pdf_path = Path(pdf_path)
    result = ExtractedPDF(file_path=pdf_path, file_name=pdf_path.name, page_number=page_number)
    try:
        if not pdf_path.exists():
            raise FileNotFoundError(str(pdf_path))
        doc = fitz.open(pdf_path)
        try:
            if page_number < 1 or page_number > len(doc):
                raise ValueError(f"PDF has {len(doc)} pages; requested page {page_number}.")
            page = doc[page_number - 1]
            # PyMuPDF may print an optional package recommendation during table detection.
            # Suppress it so CLI/GUI output stays focused on validation results.
            with contextlib.redirect_stdout(io.StringIO()), contextlib.redirect_stderr(io.StringIO()):
                found = page.find_tables()
            tables = list(found.tables or [])
            if not tables:
                raise ValueError("No table detected on the configured page. Is the page number correct?")

            candidates: List[Tuple[float, int, Sequence[Sequence[object]]]] = []
            selected_tables = []
            if table_index is not None:
                if table_index < 1 or table_index > len(tables):
                    raise ValueError(f"Table index {table_index} invalid. Detected {len(tables)} table(s) on the page.")
                selected_tables = [(table_index, tables[table_index - 1])]
            else:
                selected_tables = [(i + 1, t) for i, t in enumerate(tables)]

            for idx, table in selected_tables:
                rows = table.extract()
                if not rows:
                    continue
                if row_start is not None:
                    score = _candidate_score_from_rows(rows, row_start - 1)
                    candidates.append((score, idx, rows))
                else:
                    max_score = max((_candidate_score_from_rows(rows, s) for s in range(0, max(0, len(rows) - 8))), default=-1)
                    # Slightly prefer larger tables because the metadata table has many rows.
                    max_score += min(len(rows), 20) * 0.1
                    candidates.append((max_score, idx, rows))

            if not candidates:
                raise ValueError("No readable table rows found on the configured page.")
            _score, chosen_idx, rows = max(candidates, key=lambda x: x[0])
            return _extract_from_table_rows(rows, pdf_path, page_number, chosen_idx, row_start)
        finally:
            doc.close()
    except Exception as exc:
        result.error = str(exc)
        return result


def compare_value(field_name: str, pdf_raw: str, reference_raw: str, source: str, file_name: str, excel_row: Optional[int]) -> FieldComparison:
    pdf_norm = normalize_field(field_name, pdf_raw)
    ref_norm = normalize_field(field_name, reference_raw)
    if field_name == "ausgabe":
        ref_norm = normalize_expected_ausgabe(reference_raw) if re.fullmatch(r"\d{1,2}\.\d{4}", clean_text(reference_raw)) else normalize_month_year(reference_raw)

    if not pdf_raw and not reference_raw:
        res = "BOTH_EMPTY"
    elif not pdf_raw:
        res = "PDF_MISSING"
    elif not reference_raw:
        res = "REFERENCE_MISSING"
    elif field_name == "ausgabe" and not pdf_norm:
        res = "PDF_UNSUPPORTED_DATE_FORMAT"
    elif pdf_norm == ref_norm and clean_text(pdf_raw) == clean_text(reference_raw):
        res = "OK"
    elif pdf_norm == ref_norm:
        res = "OK_NORMALIZED"
    else:
        res = "MISMATCH"
    return FieldComparison(
        file_name=file_name,
        field=PDF_FIELD_DISPLAY[field_name],
        source=source,
        pdf_raw=clean_text(pdf_raw),
        pdf_norm=pdf_norm,
        reference_raw=clean_text(reference_raw),
        reference_norm=ref_norm,
        result=res,
        excel_row=excel_row,
    )


def score_candidate(pdf: ExtractedPDF, record: ExcelRecord) -> float:
    earned = 0
    possible = 0
    for field_name, weight in MATCH_WEIGHTS.items():
        pdf_norm = pdf.values_norm.get(field_name, "")
        xl_norm = record.values_norm.get(field_name, "")
        if not pdf_norm or not xl_norm:
            continue
        possible += weight
        if pdf_norm == xl_norm:
            earned += weight
    if possible == 0:
        return 0.0
    return round(100.0 * earned / possible, 1)


def build_dok_id_index(records: Iterable[ExcelRecord]) -> Dict[str, List[ExcelRecord]]:
    idx: Dict[str, List[ExcelRecord]] = {}
    for rec in records:
        dok = rec.values_norm.get("dok_id", "")
        if dok:
            idx.setdefault(dok, []).append(rec)
    return idx


def _candidate_list(pdf: ExtractedPDF, records: Sequence[ExcelRecord]) -> List[Tuple[ExcelRecord, float]]:
    scored = [(rec, score_candidate(pdf, rec)) for rec in records]
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored


def _best_candidate_is_clear(scored: List[Tuple[ExcelRecord, float]], threshold: float = 70.0, margin: float = 15.0) -> bool:
    if not scored:
        return False
    best = scored[0][1]
    second = scored[1][1] if len(scored) > 1 else 0
    return best >= threshold and (best - second >= margin or best >= 90.0)


def validate_one_pdf(
    extracted: ExtractedPDF,
    records: Sequence[ExcelRecord],
    dok_id_index: Dict[str, List[ExcelRecord]],
    expected_ausgabe_input: str,
    expected_ausgabe_norm: str,
) -> ValidationResult:
    if extracted.error:
        return ValidationResult(
            extracted=extracted,
            status="EXTRACTION_ERROR",
            issues=[extracted.error],
        )

    result = ValidationResult(extracted=extracted, status="NO_RELIABLE_MATCH")

    # Ausgabe is always checked against user input, independently of Excel.
    ausgabe_comp = FieldComparison(
        file_name=extracted.file_name,
        field=PDF_FIELD_DISPLAY["ausgabe"],
        source="User input",
        pdf_raw=extracted.values_raw.get("ausgabe", ""),
        pdf_norm=extracted.values_norm.get("ausgabe", ""),
        reference_raw=expected_ausgabe_input,
        reference_norm=expected_ausgabe_norm,
        result="OK" if extracted.values_norm.get("ausgabe", "") == expected_ausgabe_norm and clean_text(extracted.values_raw.get("ausgabe", "")) == expected_ausgabe_input else (
            "OK_NORMALIZED" if extracted.values_norm.get("ausgabe", "") == expected_ausgabe_norm else (
                "PDF_UNSUPPORTED_DATE_FORMAT" if extracted.values_raw.get("ausgabe") and not extracted.values_norm.get("ausgabe") else "MISMATCH"
            )
        ),
    )

    pdf_dok = extracted.values_norm.get("dok_id", "")
    matched_record: Optional[ExcelRecord] = None
    match_reason = ""

    if pdf_dok and pdf_dok in dok_id_index:
        rows = dok_id_index[pdf_dok]
        if len(rows) == 1:
            matched_record = rows[0]
            result.confidence = 100.0
            match_reason = "DOK-ID found exactly once in Excel."
        else:
            scored = _candidate_list(extracted, rows)
            result.candidate_rows = [(rec.row_number, score) for rec, score in scored[:10]]
            if _best_candidate_is_clear(scored, threshold=70.0, margin=10.0):
                matched_record = scored[0][0]
                result.confidence = scored[0][1]
                result.status = "DUPLICATE_DOK_ID_RESOLVED"
                match_reason = f"DOK-ID appears {len(rows)} times; best duplicate row selected."
                result.issues.append(match_reason)
            else:
                result.status = "DUPLICATE_DOK_ID_AMBIGUOUS"
                result.confidence = scored[0][1] if scored else 0.0
                result.issues.append(f"DOK-ID appears {len(rows)} times and no unique best row was found.")
                result.comparisons.append(ausgabe_comp)
                return result
    else:
        scored = _candidate_list(extracted, records)
        result.candidate_rows = [(rec.row_number, score) for rec, score in scored[:10]]
        if _best_candidate_is_clear(scored, threshold=70.0, margin=15.0):
            matched_record = scored[0][0]
            result.confidence = scored[0][1]
            result.status = "LIKELY_WRONG_DOK_ID"
            issue = "DOK-ID from PDF is empty or not present in Excel, but another row strongly matches other fields."
            result.issues.append(issue)
            match_reason = issue
        else:
            result.confidence = scored[0][1] if scored else 0.0
            result.status = "DOK_ID_NOT_FOUND" if pdf_dok else "NO_RELIABLE_MATCH"
            if pdf_dok:
                result.issues.append("DOK-ID from PDF was not found in Excel and no reliable fallback match was found.")
            else:
                result.issues.append("DOK-ID in PDF is empty and no reliable fallback match was found.")
            result.comparisons.append(ausgabe_comp)
            return result

    result.matched_excel_row = matched_record.row_number if matched_record else None

    # Compare PDF fields with Excel row.
    comps: List[FieldComparison] = []
    if matched_record:
        for field_name in ["sprache", "dokumentnummer", "dok_id", "freigabe", "artikelnummer", "revision", "version"]:
            comp = compare_value(
                field_name,
                extracted.values_raw.get(field_name, ""),
                matched_record.values_raw.get(field_name, ""),
                source="Excel",
                file_name=extracted.file_name,
                excel_row=matched_record.row_number,
            )
            comps.append(comp)
    comps.append(ausgabe_comp)
    result.comparisons = comps

    excel_bad = [c for c in comps if c.source == "Excel" and c.result not in {"OK", "OK_NORMALIZED", "BOTH_EMPTY"}]
    normalized_only = any(c.result == "OK_NORMALIZED" for c in comps)
    ausgabe_bad = ausgabe_comp.result not in {"OK", "OK_NORMALIZED"}

    if result.status == "DUPLICATE_DOK_ID_RESOLVED":
        if excel_bad:
            result.issues.append("Best duplicate row still has mismatching Excel fields: " + ", ".join(c.field for c in excel_bad))
        if ausgabe_bad:
            result.issues.append("Ausgabe differs from expected input.")
        return result

    if result.status == "LIKELY_WRONG_DOK_ID":
        if excel_bad:
            result.issues.append("Fallback row found, but some Excel fields still differ: " + ", ".join(c.field for c in excel_bad))
        if ausgabe_bad:
            result.issues.append("Ausgabe differs from expected input.")
        return result

    if excel_bad and ausgabe_bad:
        result.status = "MISMATCH_AND_AUSGABE_MISMATCH"
        result.issues.append("Excel fields differ and Ausgabe differs from expected input.")
    elif excel_bad:
        result.status = "MISMATCH"
        result.issues.append("Mismatching Excel fields: " + ", ".join(c.field for c in excel_bad))
    elif ausgabe_bad:
        result.status = "AUSGABE_MISMATCH"
        result.issues.append("Ausgabe differs from expected input.")
    elif normalized_only:
        result.status = "OK_WITH_NORMALIZATION"
        result.issues.append("All checks passed after normalization.")
    else:
        result.status = "OK"
        if match_reason:
            result.issues.append(match_reason)
    return result


def list_pdf_files(pdf_folder: Path | str) -> List[Path]:
    folder = Path(pdf_folder)
    if not folder.exists() or not folder.is_dir():
        raise NotADirectoryError(f"PDF folder not found: {folder}")
    return sorted([p for p in folder.iterdir() if p.is_file() and p.suffix.casefold() == ".pdf"])


def run_validation(
    pdf_folder: Path | str,
    excel_path: Path | str,
    output_path: Path | str,
    expected_ausgabe: str,
    sheet_name: Optional[str] = None,
    page_number: int = 2,
    table_index: Optional[int] = None,
    row_start: Optional[int] = None,
    progress_callback=None,
) -> List[ValidationResult]:
    expected_norm = normalize_expected_ausgabe(expected_ausgabe)
    records, actual_sheet, header_row, columns = load_excel_records(excel_path, sheet_name)
    if not records:
        raise ValueError("No data rows found in Excel after the required header row.")
    dok_idx = build_dok_id_index(records)
    pdfs = list_pdf_files(pdf_folder)
    if not pdfs:
        raise ValueError("No PDF files found in the selected folder.")

    results: List[ValidationResult] = []
    for i, pdf in enumerate(pdfs, start=1):
        if progress_callback:
            progress_callback(f"Processing {i}/{len(pdfs)}: {pdf.name}")
        extracted = extract_pdf_table(pdf, page_number=page_number, table_index=table_index, row_start=row_start)
        res = validate_one_pdf(extracted, records, dok_idx, expected_ausgabe, expected_norm)
        results.append(res)

    if progress_callback:
        progress_callback("Writing report...")
    write_report(results, records, output_path, expected_ausgabe, expected_norm, actual_sheet)
    if progress_callback:
        progress_callback(f"Done. Report saved to: {output_path}")
    return results


def preview_extraction(
    pdf_folder: Path | str,
    page_number: int = 2,
    table_index: Optional[int] = None,
    row_start: Optional[int] = None,
    limit: int = 5,
) -> List[ExtractedPDF]:
    pdfs = list_pdf_files(pdf_folder)[:limit]
    return [extract_pdf_table(pdf, page_number=page_number, table_index=table_index, row_start=row_start) for pdf in pdfs]


def _status_fill(status: str) -> PatternFill:
    colors = {
        "OK": "C6EFCE",
        "OK_WITH_NORMALIZATION": "D9EAD3",
        "AUSGABE_MISMATCH": "FCE4D6",
        "MISMATCH": "FFC7CE",
        "MISMATCH_AND_AUSGABE_MISMATCH": "FF9999",
        "LIKELY_WRONG_DOK_ID": "FFF2CC",
        "DUPLICATE_DOK_ID_RESOLVED": "DDEBF7",
        "DUPLICATE_DOK_ID_AMBIGUOUS": "F4CCCC",
        "DOK_ID_NOT_FOUND": "F4CCCC",
        "NO_RELIABLE_MATCH": "F4CCCC",
        "EXTRACTION_ERROR": "B4C6E7",
    }
    return PatternFill("solid", fgColor=colors.get(status, "FFFFFF"))


def _style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = border
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for col in range(1, ws.max_column + 1):
        letter = get_column_letter(col)
        max_len = 10
        for cell in ws[letter]:
            val = clean_text(cell.value)
            if val:
                max_len = max(max_len, min(len(val), 60))
        ws.column_dimensions[letter].width = min(max_len + 2, 45)


def write_report(
    results: Sequence[ValidationResult],
    excel_records: Sequence[ExcelRecord],
    output_path: Path | str,
    expected_ausgabe: str,
    expected_ausgabe_norm: str,
    sheet_name: str,
):
    output_path = Path(output_path)
    wb = Workbook()
    ws = wb.active
    ws.title = "Summary"
    ws.append([
        "PDF file",
        "Status",
        "Confidence",
        "Excel row",
        "PDF DOK-ID",
        "Expected Ausgabe",
        "PDF Ausgabe raw",
        "PDF Ausgabe normalized",
        "Detected table",
        "Row start",
        "Issues",
    ])
    for res in sorted(results, key=lambda r: (STATUS_ORDER.get(r.status, 99), r.extracted.file_name.casefold())):
        ext = res.extracted
        ws.append([
            ext.file_name,
            res.status,
            res.confidence,
            res.matched_excel_row or "",
            ext.values_raw.get("dok_id", ""),
            expected_ausgabe,
            ext.values_raw.get("ausgabe", ""),
            ext.values_norm.get("ausgabe", ""),
            ext.table_index or "",
            ext.row_start or "",
            "; ".join(res.issues + ext.warnings),
        ])
        ws.cell(ws.max_row, 2).fill = _status_fill(res.status)

    details = wb.create_sheet("Detailed comparison")
    details.append([
        "PDF file", "Status", "Excel row", "Field", "Source", "PDF raw", "PDF normalized",
        "Excel/User raw", "Excel/User normalized", "Result"
    ])
    for res in results:
        if not res.comparisons and res.extracted.error:
            details.append([res.extracted.file_name, res.status, "", "Extraction", "PDF", "", "", "", "", res.extracted.error])
            continue
        for comp in res.comparisons:
            details.append([
                comp.file_name,
                res.status,
                comp.excel_row or "",
                comp.field,
                comp.source,
                comp.pdf_raw,
                comp.pdf_norm,
                comp.reference_raw,
                comp.reference_norm,
                comp.result,
            ])
            if comp.result not in {"OK", "OK_NORMALIZED", "BOTH_EMPTY"}:
                details.cell(details.max_row, 10).fill = PatternFill("solid", fgColor="FFC7CE")
            elif comp.result == "OK_NORMALIZED":
                details.cell(details.max_row, 10).fill = PatternFill("solid", fgColor="D9EAD3")

    extraction = wb.create_sheet("Extracted PDF values")
    extraction.append(["PDF file", "Page", "Table", "Row start", "Field", "Raw value", "Normalized value", "Warnings/Error"])
    for res in results:
        ext = res.extracted
        if ext.error:
            extraction.append([ext.file_name, ext.page_number, "", "", "ERROR", "", "", ext.error])
            continue
        for field_name in PDF_FIELDS:
            extraction.append([
                ext.file_name,
                ext.page_number,
                ext.table_index,
                ext.row_start,
                PDF_FIELD_DISPLAY[field_name],
                ext.values_raw.get(field_name, ""),
                ext.values_norm.get(field_name, ""),
                "; ".join(ext.warnings),
            ])

    candidates = wb.create_sheet("Candidate rows")
    candidates.append(["PDF file", "Status", "Candidate Excel row", "Candidate score"])
    for res in results:
        for row_num, score in res.candidate_rows[:10]:
            candidates.append([res.extracted.file_name, res.status, row_num, score])

    matched_rows = {res.matched_excel_row for res in results if res.matched_excel_row}
    unmatched = wb.create_sheet("Excel rows not matched")
    unmatched.append(["Excel row", "DOK-ID", "Dok.-Nr.", "Artikel-Nr.", "Freigabe-/ Änd.-Nr.", "Rev.", "Vers.", "Sprache"])
    for rec in excel_records:
        if rec.row_number not in matched_rows:
            unmatched.append([
                rec.row_number,
                rec.values_raw.get("dok_id", ""),
                rec.values_raw.get("dokumentnummer", ""),
                rec.values_raw.get("artikelnummer", ""),
                rec.values_raw.get("freigabe", ""),
                rec.values_raw.get("revision", ""),
                rec.values_raw.get("version", ""),
                rec.values_raw.get("sprache", ""),
            ])

    meta = wb.create_sheet("Run info")
    meta.append(["Property", "Value"])
    meta.append(["Expected Ausgabe input", expected_ausgabe])
    meta.append(["Expected Ausgabe normalized", expected_ausgabe_norm])
    meta.append(["Excel sheet", sheet_name])
    meta.append(["PDF files processed", len(results)])
    meta.append(["Excel records loaded", len(excel_records)])

    for sheet in wb.worksheets:
        _style_sheet(sheet)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)


def summarize_results(results: Sequence[ValidationResult]) -> Dict[str, int]:
    counts: Dict[str, int] = {}
    for res in results:
        counts[res.status] = counts.get(res.status, 0) + 1
    return dict(sorted(counts.items(), key=lambda kv: STATUS_ORDER.get(kv[0], 99)))
