# Test results — Profile Docs Checker 1.2.0

Tested locally with generated, non-confidential files.

## PDF/Excel regression test

Passed:

- selectable PDF extraction from visible page 2
- multilingual PDF labels ignored through row-position extraction
- language normalization such as `de_DE` → `de`
- Ausgabe forms `07.2026`, `07-2026`, `07/2026`, year-first, and month names
- exact DOK-ID match
- mismatching field report
- likely wrong DOK-ID fallback
- duplicate DOK-ID resolution
- Ausgabe mismatch
- Excel report generation

## Existing custom Word-mapping regression test

Passed:

- Word inspector JSON generation
- custom mapping loading
- custom Word row extraction
- DOK-ID-based matching and wrong-DOK-ID handling
- Word report sheets

## Automatic change-notice Word-template test

A synthetic `.docx` reproducing the shared template structure was generated with:

- 13 logical columns
- merged bilingual title and header cells
- top-right Number
- separate old/new drawing revision columns
- `Dokument-Nr.`
- `Rev neu/new` and `Vers neu/new` under the document group

Passed:

- automatic template recognition without JSON
- correct detection of columns 0, 3, 7, and 8
- correct exclusion of the drawing-only `Rev neu/new` column
- top-right Number extraction
- wrapped/merged Word-header handling
- revision/version normalization
- Word Number vs Excel Freigabe-/Änd.-Nr. mismatch reporting
- `Word template info` report sheet

Generated sample report: `test_output/validation_report_auto_word.xlsx`.
