from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import write_word_inspection_files


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspect DOCX tables locally. The standard change-notice template is recognized automatically; the JSON mapping is for other Word templates."
    )
    parser.add_argument("docx", help="Word .docx file to inspect.")
    parser.add_argument("--structure-out", default="word_table_structure.json", help="Output JSON describing detected tables. Default: word_table_structure.json")
    parser.add_argument("--mapping-out", default="word_mapping_template.json", help="Output editable mapping JSON. Default: word_mapping_template.json")
    parser.add_argument("--include-samples", action="store_true", help="Include actual sample row values in the structure JSON. Default redacts data cells.")
    args = parser.parse_args()

    structure_out, mapping_out = write_word_inspection_files(
        Path(args.docx),
        structure_out=Path(args.structure_out),
        mapping_out=Path(args.mapping_out),
        include_samples=args.include_samples,
    )
    print(f"Saved Word structure: {structure_out}")
    print(f"Saved mapping template: {mapping_out}")
    print("For the standard Freigabe-/Änderungsmitteilung template, no mapping is required. Use the mapping only for another table template.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
