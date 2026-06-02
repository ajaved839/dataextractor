from __future__ import annotations

import argparse
from pathlib import Path

from .exporter import write_results
from .extractor import extract_workbook
from .paths import default_output_path


def main() -> int:
    parser = argparse.ArgumentParser(description="Extract discharge summary XLSX files into one workbook.")
    parser.add_argument("files", nargs="+", help="Input .xlsx files")
    parser.add_argument(
        "-o",
        "--output",
        default=str(default_output_path()),
        help="Output .xlsx path. Defaults to Desktop/extracted_discharge_data.xlsx.",
    )
    parser.add_argument(
        "--append-to",
        metavar="EXISTING",
        help="Existing extracted .xlsx file. New rows are appended underneath.",
    )
    args = parser.parse_args()

    results = [extract_workbook(Path(filename)) for filename in args.files]
    write_results(results, args.output, existing_path=args.append_to)

    for result in results:
        if result.missing_fields:
            print(f"{result.source_file}: missing {', '.join(result.missing_fields)}")
        else:
            print(f"{result.source_file}: ok")
    print(f"Saved {len(results)} row(s) to {args.output}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
