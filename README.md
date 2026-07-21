# Profile Docs Checker 1.5.1

Profile Docs Checker validates document metadata across a fixed Excel master list, selectable PDFs, and the standard bilingual Word **Freigabe-/Änderungsmitteilung / Engineering Change Notice** template.

## What version 1.5 changes

- The Word **Freigabe-/Änd.-Nr.** is read from the confirmed address `section[0].header.table[0].row[0].cell[1].paragraph[0].word[0]`. This exact header location takes precedence over any Number-like value in the body table.
- Text nested inside Word content controls or split across formatted runs is read correctly.
- The right side of the GUI is now the concise live report:
  - green confirmation when every selected check passes;
  - otherwise, every incorrect or unverified item is highlighted directly;
  - each line shows the reliable Excel DOK-ID when available and only the incorrect information beside it.
- The old scrolling technical log is replaced with a small current-step indicator.
- A new **Create Corrected Word Copy** action produces a separate `.docx` without modifying the original.
- The corrected copy uses:
  - Excel values for `Artikel-Nr.`, `Dokument-Nr.`, document `Rev neu`, document `Vers neu`, and the top-right `Number`;
  - actual PDF page counts for `Blatt / sheets` when PDFs are supplied.
- Ambiguous or unmatched Word rows are skipped rather than guessed.

## Comparison rules

Technical values are compared character-for-character after removing only layout artifacts such as line breaks and surrounding whitespace. Prefixes, asterisks, hyphens, case, punctuation and leading zeros remain meaningful.

Only these fields use explicit conversion rules:

- PDF language-country code to the Excel language code;
- Ausgabe month/year to a common internal month/year;
- Word `Blatt / sheets` to an integer page count.

Japanese uses the local code `jp`; both `ja_JP` and `jp_JP` compare with Excel `jp`.

## Validation modes

### PDF + Excel

The tool opens visible PDF page `2`, extracts the fixed metadata table by row position and compares:

- Sprache → Excel `Sprache`
- Dokumentnummer → `Dok.-Nr.`
- DOK-ID → `DOK-ID`
- Freigabe-/Änd.-Nr. → `Freigabe-/ Änd.-Nr.`
- Artikelnummer → `Artikel-Nr.`
- Revision → `Rev.`
- Version → `Vers.`

`Dokumentname` is ignored. Expected Ausgabe accepts `07.2026`, `07-2026`, `07/2026`, or `07 2026`.

### Word + Excel

The standard 13-column change-notice template is recognized automatically. In scope:

- dedicated top-right `Number` → Excel `Freigabe-/ Änd.-Nr.`
- `Artikel-Nr. / part no.` → Excel `Artikel-Nr.`
- `Dokument-Nr. / document no.` → Excel `Dok.-Nr.`
- document-group `Rev neu/new` → Excel `Rev.`
- document-group `Vers neu/new` → Excel `Vers.`
- `Blatt / sheets` → actual PDF page count when PDFs are supplied

Drawing-only columns, old values, Format, disposal, approval relevance and description are ignored.

### Word + Excel without PDFs

Supported. Word values are compared with Excel, but page counts are shown as unverified because the actual PDF page total is unavailable.

## On-screen result

The right panel shows the complete concise outcome. It does not reproduce the detailed workbook.

- If everything is correct: **ALL ENTRIES ARE CORRECT**.
- Otherwise: one highlighted row per incorrect or unverified item.
- `ID-NR.` contains the reliable Excel DOK-ID. It remains blank when the DOK-ID itself is the incorrect value.

## Corrected Word copy

After a successful run with a Word file, click **Create Corrected Word Copy**.

The tool saves a new document and leaves the original untouched. It updates only rows that were matched reliably to Excel. The top-right Number is changed only when all matched Excel rows agree on the same Freigabe-/Änd.-Nr.; otherwise it is left unchanged and a warning is shown.

## Excel report

The generated workbook still contains only:

- `Overview`
- `Issues`

## Install or update from GitHub

```powershell
py -m pip install --upgrade --force-reinstall --no-cache-dir "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

Check the installed version:

```powershell
py -c "import importlib.metadata as m; print(m.version('profile-docs-checker'))"
```

It should print `1.5.1`.

## Run the GUI

```powershell
profile-docs-checker-gui
```

When the Windows Scripts folder is not on PATH:

```powershell
py -c "from doc_table_checker.gui import main; main()"
```

## Command-line examples

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026
```

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx" --ausgabe 07.2026
```

```powershell
profile-docs-checker --excel "C:\master.xlsx" --word "C:\docs\change-notice.docx" --out "C:\report.xlsx"
```

## Notes

- Word input must be `.docx`, not legacy `.doc`.
- Keep source files on a local disk while running.
- Close the output report in Excel before overwriting it.
- Do not upload confidential source documents to a public repository.
