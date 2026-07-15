# Profile Docs Checker 1.3.0

Profile Docs Checker validates document metadata across a fixed Excel master list, selectable PDFs, and the standard bilingual Word **Freigabe-/Änderungsmitteilung / Engineering Change Notice** template.

## What version 1.3 adds

- Reads the true page count from every PDF.
- Reads `Blatt / sheets` from each document row in the Word template.
- Links Word rows and PDFs through their matched Excel row, then compares the Word sheet count with the actual PDF page count.
- Can run **Word vs Excel without PDFs**. The GUI displays a warning and the report records that page-count verification was unavailable.
- Word comparison starts automatically when a Word file is selected; there is no checkbox.
- The standard Word template no longer needs a mapping JSON or preview step.
- The GUI was rebuilt as a cleaner dark control-panel interface with cyan accents.

## Validation modes

### PDF + Excel

For each PDF, the tool opens visible page `2`, extracts the fixed metadata table by row position, and compares:

- Sprache → Excel `Sprache`
- Dokumentnummer → `Dok.-Nr.`
- DOK-ID → `DOK-ID`
- Freigabe-/Änd.-Nr. → `Freigabe-/ Änd.-Nr.`
- Artikelnummer → `Artikel-Nr.`
- Revision → `Rev.`
- Version → `Vers.`

`Dokumentname` is ignored. The expected Ausgabe accepts `07.2026`, `07-2026`, `07/2026`, or `07 2026`.

### Word + Excel

The tool automatically reads the standard 13-column change-notice template:

- top-right `Number` → Excel `Freigabe-/ Änd.-Nr.`
- `Artikel-Nr. / part no.` → Excel `Artikel-Nr.`
- `Dokument-Nr. / document no.` → Excel `Dok.-Nr.`
- document-group `Rev neu/new` → Excel `Rev.`
- document-group `Vers neu/new` → Excel `Vers.`
- `Blatt / sheets` → actual PDF page count when PDFs are supplied

Drawing-only revision columns, old values, Format, disposal, approval relevance, and description are ignored.

### Word + Excel without PDFs

This mode is supported. Word rows are compared with Excel normally, while every page-count result is marked `UNAVAILABLE_NO_PDFS`. The GUI asks for confirmation before proceeding.

## Install or update from GitHub

```powershell
py -m pip install --upgrade --force-reinstall --no-cache-dir "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

Check the installed version:

```powershell
py -c "import importlib.metadata as m; print(m.version('profile-docs-checker'))"
```

It should print `1.3.0`.

## Run the GUI

```powershell
profile-docs-checker-gui
```

When the Windows Scripts folder is not on PATH, use:

```powershell
py -c "from doc_table_checker.gui import main; main()"
```

### GUI workflow

1. Select the Excel master list.
2. Select a Word file, a PDF folder, or both.
3. Enter Ausgabe only when PDFs are selected.
4. Choose the output report.
5. Click **Run Validation**.

Selecting a Word file automatically enables Word checking. Selecting PDFs enables PDF metadata checks and true page-count verification.

## Command-line examples

PDF and Excel:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026
```

PDF, Excel, and Word:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx" --ausgabe 07.2026
```

Word and Excel without PDFs:

```powershell
profile-docs-checker --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx"
```

## Report sheets

Depending on selected inputs, the report contains:

- `Summary`
- `Detailed comparison`
- `Extracted PDF values`
- `Candidate rows`
- `Excel rows not matched`
- `Word vs Excel Summary`
- `Word vs Excel Details`
- `Word page count`
- `Extracted Word values`
- `Word template info`
- `Word candidate rows`
- `Excel rows not in Word`
- `Run info`

The `Word page count` sheet shows the Word value, normalized page count, matched PDF, actual PDF pages, result, and explanation.

## Notes

- Word input must be `.docx`, not legacy `.doc`.
- Keep source files on a local disk while running.
- Close the output report in Excel before overwriting it.
- Do not upload real client PDFs, Excel files, or Word files to a public repository.
