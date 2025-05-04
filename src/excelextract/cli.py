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
    parser.add_argument("config", type=Path, help="Path to the JSON configuration file.")
    parser.add_argument("-v", "--verbose", action="store_true", help="Enable verbose output.")
    parser.add_argument("-i", "--input", type=Path, help="Input folder containing Excel files, overrides config.")
    parser.add_argument("-o", "--output", type=Path, help="Output folder for CSV files, overrides config.")
    parser.add_argument("-r", "--regex", type=str, help="Regex to filter input files, overrides config.")

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
        print("Warning: No exports defined in the configuration.")

    for exportConfig in exports:
        if args.input:
            exportConfig["inputFolder"] = str(args.input)
        if args.output:
            exportConfig["outputFolder"] = str(args.output)
        if "output" not in exportConfig:
            exportConfig["output"] = "output.csv"
        if args.regex:
            exportConfig["inputRegex"] = args.regex
        if "inputRegex" not in exportConfig:
            exportConfig["inputRegex"] = ".*"

        extract(exportConfig)

    print("Processing completed.")

if __name__ == "__main__":
    main()
