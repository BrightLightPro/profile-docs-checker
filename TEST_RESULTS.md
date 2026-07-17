# Test results — Profile Docs Checker 1.5.0

Tested locally with generated, non-confidential PDF, Excel and Word fixtures.

## Regression tests

Passed:

- selectable PDF extraction from visible page 2
- multilingual fixed-position PDF table extraction
- strict preservation of values including `EE******` and leading zeros
- language conversion including Japanese `ja_JP` / `jp_JP` → `jp`
- Ausgabe comparison
- missing and duplicate DOK-ID strategies
- standard Word-template recognition
- Word-only validation
- actual PDF page-count comparison
- two-sheet Excel report generation

## Top-right Number fix

Passed with a Word fixture containing:

- the change-notice title in the wide top-left cell;
- `(Number) FR-1` in the dedicated top-right cell;
- a misleading code `DISTRACTOR-999` elsewhere in the header.

The extractor returned `FR-1` from row `0`, column `12` and ignored the distractor. The XML-level reader also supports text nested in content controls and text split across formatted runs.

## Concise in-app summary model

Passed:

- all incorrect Word/PDF fields become concise issue rows;
- a reliable Excel DOK-ID is attached when available;
- the ID is intentionally omitted when the DOK-ID itself is wrong;
- unverified page counts without PDFs become warnings;
- all-correct runs produce no issue rows.

## Corrected Word copy

Passed:

- original Word document remains untouched;
- top-right Number corrected from Excel;
- Artikel-Nr., Dokument-Nr., document Rev neu and Vers neu corrected from Excel;
- Blatt / sheets corrected from actual PDF page count;
- unmatched rows are skipped rather than guessed;
- the corrected document reopens and is re-extracted with the expected values.

The generated corrected fixture was rendered before and after correction. Table structure and page layout were preserved; only intended text values changed.

## GUI smoke test

Passed under a virtual X display:

- GUI creates successfully;
- concise results Treeview is present with `ID-NR.` and issue columns;
- corrected-Word action is present and initially disabled;
- window minimum size is sufficient to show all controls;
- no old technical log or preview/mapping controls are present.

## Installation test

Passed:

- wheel/source installation as `profile-docs-checker==1.5.0`;
- package imports successfully;
- console entry points are generated;
- installed synthetic validation creates the two-sheet report.
