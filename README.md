# Profile Docs Checker

A small Windows-friendly Python tool for validating fixed-template PDF document-identification tables against an Excel master list. Version 1.1.0 also optionally compares a document list table in a Word `.docx` file against the same Excel file.

## What it checks

### PDF vs Excel

For every PDF in a selected folder, the tool opens the configured page, default visible page `2`, extracts the fixed metadata table by row position, and compares these fields with Excel:

- Sprache
- Dokumentnummer / `Dok.-Nr.`
- DOK-ID
- Freigabe-/Änd.-Nr.
- Artikelnummer / `Artikel-Nr.`
- Revision / `Rev.`
- Version / `Vers.`

`Dokumentname` is ignored.

### Ausgabe check

The expected Ausgabe is entered once globally before the run. Preferred format is:

```text
07.2026
```

The tool also accepts:

```text
07-2026
07/2026
07 2026
```

Internally all of these are normalized to `YYYY-MM`.

### Language normalization

The PDF language value can be like `de_DE` or `en_US`. Excel can contain just `de` or `en`. The tool compares only the language part, so `en_US` and `en_GB` both match Excel `en`.

### Word vs Excel, optional

If enabled, the tool reads a table in a Word `.docx` file using a local JSON mapping file, then compares the Word table rows against the Excel list using the same matching strategy:

1. match by DOK-ID
2. resolve duplicate DOK-ID by comparing the other fields
3. fallback match if the DOK-ID is missing or wrong
4. report ambiguous/unmatched rows for review

The Word template does not need to be shared. Run the local inspector to create the mapping file.

## Install

From a local folder:

```powershell
py -m pip install --upgrade --force-reinstall "C:\path\to\profile_docs_checker_v110"
```

From the ZIP:

```powershell
py -m pip install --upgrade --force-reinstall "C:\path\to\profile_docs_checker_tool_pip_v1_1_0.zip"
```

From your GitHub repo after pushing the updated files:

```powershell
py -m pip install --upgrade --force-reinstall "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

## Run the GUI

Preferred new command:

```powershell
profile-docs-checker-gui
```

Backward-compatible old command:

```powershell
doc-table-checker-gui
```

## Run from command line

PDF/Excel only:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026
```

PDF/Excel plus Word comparison:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026 --word "C:\docs\list.docx" --word-mapping "C:\docs\word_mapping_template.json"
```

## Inspect a Word file locally

Run this once on the confidential Word file:

```powershell
profile-docs-checker-inspect-word "C:\docs\list.docx" --structure-out "C:\docs\word_table_structure.json" --mapping-out "C:\docs\word_mapping_template.json"
```

The structure JSON contains table/header information. By default, sample data cells are redacted.

The mapping JSON looks like this:

```json
{
  "table_index_zero_based": 0,
  "header_row_zero_based": 0,
  "data_start_row_zero_based": 1,
  "columns_by_index_zero_based": {
    "sprache": 0,
    "dokumentnummer": 1,
    "dok_id": 2,
    "freigabe": 3,
    "artikelnummer": 4,
    "revision": 5,
    "version": 6
  }
}
```

Column indexes are zero-based: first Word table column = `0`, second = `1`, etc. Leave a field as `null` to ignore it.

## Report sheets

The generated Excel report includes the previous PDF sheets plus optional Word sheets:

- `Summary`
- `Detailed comparison`
- `Extracted PDF values`
- `Candidate rows`
- `Excel rows not matched`
- `Word vs Excel Summary` if Word comparison is enabled
- `Word vs Excel Details` if Word comparison is enabled
- `Extracted Word values` if Word comparison is enabled
- `Word candidate rows` if Word comparison is enabled
- `Excel rows not in Word` if Word comparison is enabled
- `Run info`

## Safe workflow

1. Copy the PDFs, Excel file, and Word file to a local disk.
2. Make sure the output report is not open in Excel.
3. Run the PDF extraction preview.
4. For Word checking, run the Word inspector and review the mapping JSON.
5. Run full validation.

## Notes

- The Word feature supports `.docx`, not old `.doc` files.
- Do not upload real client PDFs, Excel files, or Word files to GitHub.
- The tool is designed for selectable/text-based PDFs, not scans.
