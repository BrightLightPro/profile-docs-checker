# Profile Docs Checker 1.2.0

A Windows-friendly Python tool that validates fixed-template PDF document-identification tables against an Excel master list. It can also compare the standard bilingual Word **Freigabe-/Änderungsmitteilung / Engineering Change Notice** table with the same Excel list.

## PDF checks

For every PDF in the selected folder, the tool opens visible page `2`, extracts the fixed metadata table by position, and compares:

- Sprache → Excel `Sprache`
- Dokumentnummer → `Dok.-Nr.`
- DOK-ID → `DOK-ID`
- Freigabe-/Änd.-Nr. → `Freigabe-/ Änd.-Nr.`
- Artikelnummer → `Artikel-Nr.`
- Revision → `Rev.`
- Version → `Vers.`

`Dokumentname` is ignored. The global Ausgabe input accepts `07.2026`, `07-2026`, `07/2026`, or `07 2026` and is compared with the Ausgabe in every PDF.

## Automatic Word-template check

For the standard Word template shown as **Freigabe-/Änderungsmitteilung / Engineering Change Notice**, no mapping JSON is needed.

The tool automatically recognizes the 13-column table and reads:

- top-right **Number** → Excel `Freigabe-/ Änd.-Nr.` for every listed row
- `Artikel-Nr. / part no.` → Excel `Artikel-Nr.`
- `Dokument-Nr. / document no.` → Excel `Dok.-Nr.`
- `Rev neu/new` under `Zeichnung / Dokument / Stückliste` → Excel `Rev.`
- `Vers neu/new` under the same group → Excel `Vers.`

It deliberately ignores the drawing-only old/new revision columns, Format, old revision/version, sheets, disposal, approval relevance, and description.

Word rows are matched primarily by Dokument-Nr. Duplicate document numbers are resolved with Artikel-Nr., new revision, and new version. If Dokument-Nr. is wrong or absent, the tool tries a cautious fallback and flags the result.

## Install or update from GitHub

```powershell
py -m pip install --upgrade --force-reinstall --no-cache-dir "git+https://github.com/brightlightpro/profile-docs-checker.git"
```

Check the installed version:

```powershell
py -c "import importlib.metadata as m; print(m.version('profile-docs-checker'))"
```

It should print `1.2.0`.

## Run the GUI

```powershell
profile-docs-checker-gui
```

The old command remains available:

```powershell
doc-table-checker-gui
```

### Word workflow in the GUI

1. Tick **Also compare Word table against Excel**.
2. Select the `.docx` change-notice file.
3. Leave **Custom mapping JSON** blank for the standard template.
4. Click **Preview Word extraction** and verify the detected columns and top-right Number.
5. Run full validation.

## Command-line examples

PDF and Excel only:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026
```

PDF, Excel, and the standard Word template:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026 --word "C:\docs\change-notice.docx"
```

For a different Word-table template, a custom mapping remains supported:

```powershell
profile-docs-checker --pdf-folder "C:\PDFs" --excel "C:\master.xlsx" --out "C:\report.xlsx" --ausgabe 07.2026 --word "C:\docs\other-list.docx" --word-mapping "C:\docs\word_mapping.json"
```

## Optional custom-table inspector

The inspector is no longer required for the standard change-notice template. It remains available for other Word tables:

```powershell
profile-docs-checker-inspect-word "C:\docs\other-list.docx" --structure-out "C:\docs\word_table_structure.json" --mapping-out "C:\docs\word_mapping_template.json"
```

## Report sheets

The report contains the PDF sheets plus these Word-specific sheets when Word checking is enabled:

- `Word vs Excel Summary`
- `Word vs Excel Details`
- `Extracted Word values`
- `Word template info`
- `Word candidate rows`
- `Excel rows not in Word`

The `Word template info` sheet records the recognized table, detected column indexes, the top-right Number, and extraction warnings.

## Important notes

- Word input must be `.docx`, not legacy `.doc`.
- Keep source files on a local disk while running.
- Close the output report in Excel before overwriting it.
- Do not upload real client PDFs, Excel files, or Word files to GitHub.
