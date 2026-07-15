# Test results — Profile Docs Checker 1.4.0

Tested locally with generated, non-confidential PDF, Excel and Word fixtures.

## Strict-value PDF/Excel tests

Passed:

- selectable PDF extraction from visible page 2
- extraction independent of PDF label language
- exact preservation and comparison of `EE******`
- exact revision/version behavior: `01` is not silently converted to `1`
- language conversion such as `de_DE` → Excel `de`
- Japanese conversion `ja_JP` and `jp_JP` → Excel `jp`
- Ausgabe comparison across dot, hyphen, slash, year-first and month-name formats
- exact, missing and duplicate DOK-ID strategies
- actual PDF page-count capture

## Compact report tests

Passed:

- exactly two sheets: `Overview` and `Issues`
- document-level Overview rows
- Issues contains only warnings, mismatches, missing records and run limitations
- raw source values are displayed without generic normalized columns
- approval-relevant and other out-of-scope Word columns are absent from the report

## Standard Word change-notice template

Passed with a synthetic 13-column `.docx` containing merged bilingual headers:

- automatic template recognition
- correct selection of Artikel-Nr., Dokument-Nr., document `Rev neu` and document `Vers neu`
- correct exclusion of drawing-only revisions and approval relevance
- top-right Number extraction
- `Blatt / sheets` detection
- Word/Excel comparison
- actual PDF page count vs Word sheets comparison
- page-count and change-number mismatch reporting

## Word-only mode

Passed:

- Excel + Word validation without a PDF folder
- no Ausgabe requirement when PDFs are absent
- one clear run-limitation warning for unavailable true page counts
- Overview page status `Unavailable — no PDFs`

## Backward compatibility

The internal custom Word-mapping functions still pass regression tests, although those controls are no longer exposed in the simplified interface.

## GUI smoke test

Passed under a virtual display:

- GUI imports and opens successfully
- light graphite/cyan control-panel appearance
- initial window size is at least 1040 × 690
- Windows launch path requests maximized state
- all main inputs, actions and the system log are present without advanced mapping/preview controls

## Package installation test

Passed:

- source package installs as `profile-docs-checker==1.4.0`
- simplified CLI help exposes only Excel, optional PDFs, optional Word, Ausgabe, sheet and output
- installed CLI completed a six-PDF validation and created the two-sheet report
