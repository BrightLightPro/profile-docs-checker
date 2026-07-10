# Test Results

Synthetic end-to-end test was run successfully with:

```bash
cd /mnt/data/doc_table_checker
python tests/test_synthetic.py
```

The test creates:

- one German-template Excel file with required columns on row 3
- five text-based PDFs with the table on page 2
- multilingual PDF labels, while extraction uses row position rather than labels
- cases for:
  - `OK_WITH_NORMALIZATION`
  - `MISMATCH`
  - `LIKELY_WRONG_DOK_ID`
  - `DUPLICATE_DOK_ID_RESOLVED`
  - `AUSGABE_MISMATCH`

Observed result counts:

```text
OK_WITH_NORMALIZATION: 1
AUSGABE_MISMATCH: 1
MISMATCH: 1
LIKELY_WRONG_DOK_ID: 1
DUPLICATE_DOK_ID_RESOLVED: 1
```

The sample generated report is in `test_output/validation_report.xlsx`.
