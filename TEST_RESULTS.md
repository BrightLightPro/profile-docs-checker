# Test Results - Profile Docs Checker v1.1.0

Tests were run with synthetic data only.

## PDF/Excel synthetic test

Command:

```bash
cd /mnt/data/profile_docs_checker_v110
python tests/test_synthetic.py
```

Result:

```text
Result counts: {'OK_WITH_NORMALIZATION': 1, 'AUSGABE_MISMATCH': 1, 'MISMATCH': 1, 'LIKELY_WRONG_DOK_ID': 1, 'DUPLICATE_DOK_ID_RESOLVED': 1}
Synthetic test passed.
Sample report: /mnt/data/profile_docs_checker_v110/test_output/validation_report.xlsx
```

Covered cases:

- normal match with language/date/revision normalization
- field mismatch
- wrong/missing DOK-ID with fallback match
- duplicated DOK-ID resolved by other fields
- Ausgabe mismatch

## Ausgabe parser test

Inputs accepted and normalized to `2026-07`:

```text
07.2026
07-2026
07/2026
7 2026
```

## Word inspector and Word/Excel comparison test

Synthetic Word document generated locally:

```text
test_output/word_list.docx
```

Inspector outputs:

```text
test_output/word_table_structure.json
test_output/word_mapping_template.json
```

Full validation with Word comparison generated:

```text
test_output/validation_report_with_word.xlsx
```

Report contains these Word-specific sheets:

```text
Word vs Excel Summary
Word vs Excel Details
Extracted Word values
Word candidate rows
Excel rows not in Word
```

Covered Word cases:

- exact Word row match by DOK-ID
- Word row with wrong DOK-ID resolved by fallback fields
- Word comparison details written to report
