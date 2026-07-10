# PDF / Excel Document Table Checker

A small Python desktop/CLI tool that validates fixed-template PDF document identification tables against a German Excel master list.

## What it checks

The tool reads **page 2** of each PDF by default and extracts values by **table row position**, not by label text. This is intentional because the PDF table may be in many languages.

PDF row logic:

| Row in PDF table | Meaning | Check |
|---:|---|---|
| 1 | Dokumentname | Ignored |
| 2 | Sprache | Compare with Excel `Sprache` after `de_DE -> de` normalization |
| 3 | Dokumentnummer | Compare with Excel `Dok.-Nr.` |
| 4 | DOK-ID | Main Excel lookup key |
| 5 | Freigabe-/Änd.-Nr. | Compare with Excel `Freigabe-/ Änd.-Nr.` |
| 6 | Artikelnummer | Compare with Excel `Artikel-Nr.` |
| 7 | Revision | Compare with Excel `Rev.` |
| 8 | Version | Compare with Excel `Vers.` |
| 9 | Ausgabe | Compare with the user-entered global expected month |

The Excel template is expected to have these German columns:

- `Sprache`
- `Dok.-Nr.`
- `DOK-ID`
- `Freigabe-/ Änd.-Nr.`
- `Artikel-Nr.`
- `Rev.`
- `Vers.`

All other Excel columns are ignored.

## Ausgabe input

Enter the global expected Ausgabe as:

```text
MM.YYYY
```

Example:

```text
07.2026
```

The PDF value may be written as `07.2026`, `07/2026`, `07-2026`, `07 2026`, `2026-07`, `2026/07`, or common month-name formats like `Juli 2026`. Internally, everything is normalized to `YYYY-MM` for comparison.

## Installation

1. Install Python 3.10 or newer.
2. Open a terminal in this folder.
3. Install dependencies:

```bash
python -m pip install -r requirements.txt
```

## Run the desktop app

```bash
python run_gui.py
```

Recommended workflow:

1. Choose the PDF folder.
2. Choose the Excel file and sheet.
3. Enter Expected Ausgabe as `MM.YYYY`.
4. Click **Preview extraction** and check that the fields are read correctly.
5. Click **Run full validation**.

## Run from command line

Preview the first five PDFs:

```bash
python validate_cli.py --pdf-folder "C:\path\to\pdfs" --excel "C:\path\to\list.xlsx" --ausgabe 07.2026 --out "C:\path\to\report.xlsx" --preview
```

Run full validation:

```bash
python validate_cli.py --pdf-folder "C:\path\to\pdfs" --excel "C:\path\to\list.xlsx" --ausgabe 07.2026 --out "C:\path\to\validation_report.xlsx"
```

Optional parameters:

```bash
--sheet "Sheet1"        # Excel sheet name. Default: first sheet.
--page 2                # visible PDF page containing the table. Default: 2.
--table-index 1         # force a specific table number on the page. Default: auto.
--row-start 1           # force the row where Dokumentname starts. Default: auto.
```

## Matching strategy

1. The tool first tries to match the PDF `DOK-ID` against Excel.
2. If the `DOK-ID` exists exactly once, that Excel row is used.
3. If the `DOK-ID` is missing from Excel, the tool scores fallback matches using:
   - `Dok.-Nr.`
   - `Artikel-Nr.`
   - `Freigabe-/ Änd.-Nr.`
   - `Rev.`
   - `Vers.`
   - `Sprache`
4. If the `DOK-ID` appears more than once, the tool compares the duplicate rows and either resolves the best row or flags it as ambiguous.

## Report sheets

The generated report contains:

- `Summary`
- `Detailed comparison`
- `Extracted PDF values`
- `Candidate rows`
- `Excel rows not matched`
- `Run info`

Important statuses:

- `OK`
- `OK_WITH_NORMALIZATION`
- `MISMATCH`
- `AUSGABE_MISMATCH`
- `LIKELY_WRONG_DOK_ID`
- `DUPLICATE_DOK_ID_RESOLVED`
- `DUPLICATE_DOK_ID_AMBIGUOUS`
- `DOK_ID_NOT_FOUND`
- `NO_RELIABLE_MATCH`
- `EXTRACTION_ERROR`

## If preview extraction is wrong

Because the PDF labels can be multilingual, the tool does not use label text. It uses table row position. If the preview looks shifted:

- Set `Row where Dokumentname starts` to `1` if the first extracted row is the Dokumentname row.
- Set `Table number on page` if the page has multiple detected tables.

Then preview again before running the full check.

## Install with pip

This project is pip-installable.

From a local folder or ZIP:

```bash
python -m pip install /path/to/doc_table_checker_tool_pip.zip
```

From a GitHub repository, after uploading the project:

```bash
python -m pip install git+https://github.com/YOUR-USER/doc-table-checker.git
```

From a GitHub Release ZIP, after uploading the ZIP as a release asset:

```bash
python -m pip install https://github.com/YOUR-USER/doc-table-checker/releases/download/v1.0.0/doc_table_checker_tool_pip.zip
```

After installation:

```bash
doc-table-checker-gui
```

Or use the command line version:

```bash
doc-table-checker --pdf-folder "C:\\path\\to\\pdfs" --excel "C:\\path\\master.xlsx" --out "C:\\path\\report.xlsx" --ausgabe 07.2026
```
