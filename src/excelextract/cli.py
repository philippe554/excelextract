#!/usr/bin/env python3

import argparse
import json
import sys
from pathlib import Path
import traceback
import glob
import importlib.metadata

from .utils import cleanConfig
from .simpleTable import resolveSimpleTable
from .io import loopFiles

def main():
    try:
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

        version = importlib.metadata.version("excelextract")

        parser.add_argument('--version', action='version', version=version)
        parser.add_argument("config", type=str, help="Path to the JSON configuration file.")
        parser.add_argument("-i", "--input", type=str, help="Input glob, overrides config.")
        parser.add_argument("-o", "--output", type=Path, help="Output folder, prefix for output files in the config.")

        args = parser.parse_args()

        configNames = glob.glob(args.config, recursive=True)
        print(f"Found {len(configNames)} config files.")

        if len(configNames) == 0:
            raise FileNotFoundError(f"No config files found matching {args.config}.")
        
        for configName in configNames:
            if not configName.endswith(".json"):
                raise ValueError(f"Config file {configName} is not a JSON file.")
            
            try:
                with open(configName, "r", encoding="utf-8") as f:
                    config = json.load(f)
            except json.JSONDecodeError as e:
                raise ValueError(f"Config file {configName} is not a valid JSON file: {e}")
            except Exception as e:
                raise RuntimeError(f"Error reading config file {configName}: {e}")
            
            config = cleanConfig(config)

            if "exports" not in config:
                raise ValueError("Config file does not contain 'exports' key.")
            exports = config["exports"]

            for exportConfig in exports:
                if args.input:
                    exportConfig["input"] = str(args.input)
                
                if "output" not in exportConfig:
                    exportConfig["output"] = "output.csv"
                if args.output:
                    exportConfig["output"] = args.output / exportConfig["output"]

                exportConfig = resolveSimpleTable(exportConfig)

                loopFiles(exportConfig)

        print("Processing completed.")

    except Exception as e:
        traceback.print_exc(file=sys.stderr)
        print()
        print(f"Error: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
