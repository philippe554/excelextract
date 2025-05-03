#!/usr/bin/env python3

import argparse
import json
import sys
from pathlib import Path

from .extract import extract

def main():
    parser = argparse.ArgumentParser(
        prog="excelextract",
        description=(
            "Extract structured CSV data from Excel (.xlsx) files using a declarative JSON configuration.\n\n"
            "This tool is designed for researchers and survey teams working with standardized Excel forms. "
            "You define what to extract via a JSON file — no programming required."
        ),
        epilog=(
            "Example usage:\n"
            "  excelextract config.json\n\n"
            "For documentation and examples, see: https://github.com/philippe554/excelextract"
        ),
        formatter_class=argparse.RawTextHelpFormatter
    )
    parser.add_argument(
        "config",
        type=Path,
        help="Path to the JSON configuration file."
    )

    args = parser.parse_args()

    if not args.config.exists():
        print(f"Error: Configuration file not found: {args.config}", file=sys.stderr)
        sys.exit(1)

    try:
        with args.config.open("r", encoding="utf-8") as f:
            config = json.load(f)
    except json.JSONDecodeError as e:
        print(f"Error: Failed to parse JSON file: {e}", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"Error: Could not read config file: {e}", file=sys.stderr)
        sys.exit(1)

    exports = config.get("exports", [])
    if not exports:
        print("Error: No exports defined in the configuration.", file=sys.stderr)
        sys.exit(1)

    for exportConfig in exports:
        extract(exportConfig)

    print("Processing completed.")

if __name__ == "__main__":
    main()
