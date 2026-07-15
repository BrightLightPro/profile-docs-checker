# Test results — Profile Docs Checker 1.3.0

Tested locally with generated, non-confidential PDF, Excel, and Word fixtures.

## PDF/Excel regression

Passed:

- selectable PDF extraction from visible page 2
- extraction independent of PDF label language
- Sprache normalization such as `de_DE` → `de`
- Ausgabe normalization for dot, hyphen, slash, year-first, and month-name formats
- exact, missing, and duplicate DOK-ID strategies
- field and Ausgabe mismatch reporting
- actual PDF page-count capture
- Excel report generation

## Standard Word change-notice template

Passed with a synthetic 13-column `.docx` containing merged bilingual headers:

- automatic template recognition without a mapping JSON
- correct selection of Artikel-Nr., Dokument-Nr., document `Rev neu`, and document `Vers neu`
- correct exclusion of the drawing-only revision columns
- top-right Number extraction
- `Blatt / sheets` detection at column 9
- Word/Excel comparison
- actual PDF page count vs Word sheets comparison
- page-count mismatch status and report details

## Word-only mode

Passed:

- Excel + Word validation with no PDF folder
- no Ausgabe requirement when PDFs are absent
- report creation with zero PDF rows
- `UNAVAILABLE_NO_PDFS` for every Word page-count check
- Run-info warning that true page counts were unavailable

## Custom Word mapping regression

Passed to preserve backward compatibility for command-line users with non-standard Word tables.

## GUI smoke and visual review

Passed under a virtual display:

- GUI imports and opens successfully
- redesigned dark graphite/cyan interface renders without overlapping controls
- removed advanced PDF controls, preview controls, Word checkbox, and Word mapping controls
- Word-only limitation warning is implemented before execution
