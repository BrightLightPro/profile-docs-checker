# Profile Docs Checker 1.6.0

Profile Docs Checker compares a fixed German Excel master list with selectable PDFs and the standard bilingual Word **Freigabe-/Änderungsmitteilung / Engineering Change Notice**. It can also fill a copy of the Word template directly from Excel.

## New in 1.6.0

- The interface is German by default. A clearly visible **Deutsch / English** selector switches the complete GUI language.
- The main action is now explicit:
  - **Dokumente prüfen und Bericht erstellen** / **Check documents and create report**
  - **Word-Datei aus Excel befüllen** / **Fill Word file from Excel**
- After selecting a Word file, choose one of two Word actions:
  - **Prüfen / Examine** compares the existing Word values with Excel.
  - **Aus Excel befüllen / Fill from Excel** creates a separate Word copy and fills one row per Excel entry.
- Fill mode writes the following values from Excel:
  - `Artikel-Nr.`
  - `Dok.-Nr.`
  - `Rev.` into document `Rev neu/new`
  - `Vers.` into document `Vers neu/new`
  - the common `Freigabe-/ Änd.-Nr.` into the confirmed top-right Word Number address
- When PDFs are selected in fill mode, `Blatt / sheets` is filled with each PDF’s actual page count. Without PDFs, that field is left blank and the app shows a warning.
- If Excel has more rows than the Word template, new formatting-preserving Word rows are appended automatically.
- The **Create corrected Word copy** button remains greyed out when the checked Word file has no correctable differences.

## Confirmed Word Number location

The top-right change number is read and written at the exact confirmed address:

```text
section[0].header.table[0].row[0].cell[1].paragraph[0].word[0]
```

In `python-docx`, the first visible word in paragraph 0 of that cell is used.

## Validation modes

### PDF + Excel

The tool reads visible PDF page 2 and compares:

- Sprache → Excel `Sprache`
- Dokumentnummer → `Dok.-Nr.`
- DOK-ID → `DOK-ID`
- Freigabe-/Änd.-Nr. → `Freigabe-/ Änd.-Nr.`
- Artikelnummer → `Artikel-Nr.`
- Revision → `Rev.`
- Version → `Vers.`

Expected Ausgabe accepts formats such as `07.2026`, `07-2026`, and `07/2026`.

### Word + Excel

The standard change-notice template is recognized automatically. In scope:

- confirmed top-right `Number` → Excel `Freigabe-/ Änd.-Nr.`
- `Artikel-Nr. / part no.` → Excel `Artikel-Nr.`
- `Dokument-Nr. / document no.` → Excel `Dok.-Nr.`
- document `Rev neu/new` → Excel `Rev.`
- document `Vers neu/new` → Excel `Vers.`
- `Blatt / sheets` → actual PDF page count when PDFs are available

Drawing-only fields, old values, Format, disposal, approval relevance and description are ignored.

## Comparison rules

Identifiers and technical values are compared exactly after removing only layout artifacts such as surrounding whitespace and line breaks. Prefixes, asterisks, punctuation, case and leading zeros remain meaningful.

Only language, Ausgabe and page count have defined conversion rules. Japanese uses the local code `jp`; both `ja_JP` and `jp_JP` compare with Excel `jp`.

## Install or update from GitHub

```powershell
py -m pip install --upgrade --force-reinstall --no-cache-dir "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

Check the installed version:

```powershell
py -c "import importlib.metadata as m; print(m.version('profile-docs-checker'))"
```

It should print `1.6.0`.

## Run the GUI

```powershell
profile-docs-checker-gui
```

When the Windows Scripts folder is not on PATH:

```powershell
py -c "from doc_table_checker.gui import main; main()"
```

## Command line

Validate:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --word "C:\change-notice.docx" --out "C:\report.xlsx" --ausgabe 07.2026
```

Fill Word from Excel:

```powershell
profile-docs-checker --fill-word --excel "C:\master.xlsx" --word "C:\blank-change-notice.docx" --out "C:\filled-change-notice.docx"
```

Add `--pdf-folder "C:\PDFs"` to fill `Blatt / sheets` with actual page counts.

## Notes

- Word input must be `.docx`, not legacy `.doc`.
- The original Word file is never overwritten.
- Keep source files on a local disk while running.
- Close the output Excel or Word file before overwriting it.
