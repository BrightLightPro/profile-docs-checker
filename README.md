# Profile Docs Checker 1.4.0

Profile Docs Checker validates document metadata across a fixed Excel master list, selectable PDFs, and the standard bilingual Word **Freigabe-/Änderungsmitteilung / Engineering Change Notice** template.

## What version 1.4 changes

- The Excel report now contains only two focused sheets: `Overview` and `Issues`.
- Identifiers and technical values are compared strictly, character-for-character after removing only layout artifacts such as line breaks and surrounding whitespace.
- Meaningful characters are preserved, including prefixes, asterisks, hyphens, case and leading zeros. For example, `EE******` remains `EE******`, and `01` is different from `1`.
- Only three fields use explicit conversion rules:
  - PDF language-country codes are compared with the Excel language code.
  - Ausgabe month/year formats are converted for comparison.
  - Word `Blatt / sheets` is interpreted as a page count and compared with the actual PDF page count.
- Japanese uses the local Excel code `jp`; both `ja_JP` and `jp_JP` from a PDF compare to `jp`.
- `Zulassungs-relevant / approval-relevant` and all other out-of-scope Word columns are ignored.
- The interface keeps the control-panel layout but uses a light graphite/cyan appearance and opens maximized on Windows.
- Preview and custom Word-mapping controls are no longer exposed in the GUI or normal CLI workflow.

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

The standard 13-column change-notice template is recognized automatically. The tool reads only:

- top-right `Number` → Excel `Freigabe-/ Änd.-Nr.`
- `Artikel-Nr. / part no.` → Excel `Artikel-Nr.`
- `Dokument-Nr. / document no.` → Excel `Dok.-Nr.`
- document-group `Rev neu/new` → Excel `Rev.`
- document-group `Vers neu/new` → Excel `Vers.`
- `Blatt / sheets` → actual PDF page count when PDFs are supplied

Drawing-only revision columns, old values, Format, disposal, approval relevance and description are ignored.

### Word + Excel without PDFs

This mode is supported. Word rows are compared with Excel normally. The GUI warns before the run, and the report records a single run limitation explaining that the true PDF page count could not be verified.

## Report sheets

### `Overview`

One compact document-level table showing:

- overall status
- Excel row and document identifiers
- linked PDF file and Word row
- PDF-vs-Excel and Word-vs-Excel status
- PDF Ausgabe and expected Ausgabe
- Word sheet count and actual PDF page count
- a concise issue summary

### `Issues`

Only warnings, mismatches, missing records and run limitations are listed. Values are shown exactly as read from each source; generic normalized-value columns are not included.

## Install or update from GitHub

```powershell
py -m pip install --upgrade --force-reinstall --no-cache-dir "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

Check the installed version:

```powershell
py -c "import importlib.metadata as m; print(m.version('profile-docs-checker'))"
```

It should print `1.4.0`.

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

PDF, Excel and Word:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx" --ausgabe 07.2026
```

Word and Excel without PDFs:

```powershell
profile-docs-checker --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx"
```

## Notes

- Word input must be `.docx`, not legacy `.doc`.
- Keep source files on a local disk while running.
- Close the output report in Excel before overwriting it.
- Do not upload real client PDFs, Excel files or Word files to a public repository.
