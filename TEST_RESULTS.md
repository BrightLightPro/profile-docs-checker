# Test results — Profile Docs Checker 1.6.0

All automated tests passed.

Covered scenarios:

- fixed page-2 PDF extraction and strict PDF/Excel comparison;
- Japanese language mapping (`ja_JP` and `jp_JP` → `jp`);
- Ausgabe parsing and mismatch reporting;
- duplicate and incorrect DOK-ID fallback behavior;
- automatic recognition of the 13-column Engineering Change Notice table;
- exact Word Number extraction and correction from `section[0].header.table[0].row[0].cell[1].paragraph[0].word[0]`;
- Word/Excel validation with and without PDFs;
- actual PDF page-count comparison with Word `Blatt / sheets`;
- corrected Word-copy generation;
- fill-from-Excel generation for a blank Word template;
- automatic addition of formatting-preserving Word rows when Excel has more entries;
- filling `Blatt / sheets` from actual PDF page counts;
- German default interface and English switch;
- clear action-button labels for validation and Word filling;
- corrected-Word button disabled for an all-correct result;
- GUI creation and language/mode switching under a virtual display;
- two-sheet Excel report regression checks.
