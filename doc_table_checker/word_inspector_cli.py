from __future__ import annotations

import argparse
from pathlib import Path

from doc_table_checker.core import write_word_inspection_files


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Inspect DOCX tables locally and create a structure JSON plus editable Word mapping template."
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
    print("Edit the mapping JSON if needed, then pass it to validation with --word-mapping.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
