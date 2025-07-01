#!/usr/bin/env python3

import argparse
import json
import sys
from pathlib import Path
import glob
import importlib.metadata
import logging

from .logger import logger, addFileHandler
from .utils import cleanConfig
from .simpleTable import resolveSimpleTable
from .io import loopFiles

def main():
    try:
        parser = argparse.ArgumentParser(
            prog="ExcelExtract",
            description=(
                "Extract structured CSV data from Excel (.xlsx) files using a declarative JSON configuration.\n"
                "You define what to extract via a JSON file — no programming required."
            ),
            epilog=(
                "Example usage:\n"
                "  excelextract config.json\n"
                "  excelextract input.xlsx --sheet Sheet1\n\n"
                "For documentation and examples, see: https://github.com/philippe554/excelextract"
            ),
            formatter_class=argparse.RawTextHelpFormatter
        )

        version = importlib.metadata.version("excelextract")

        parser.add_argument('--version', action='version', version=version)
        parser.add_argument("path", type=str, help="Path/glob to the JSON configuration files or XLSX files.")
        parser.add_argument("-i", "--input", type=str, help="Input glob, overrides config.")
        parser.add_argument("-o", "--output", type=Path, help="Output folder, prefix for output files in the config, or output file name if not using a config file.")
        parser.add_argument("-s", "--sheet", type=str, help="Sheet name to process, when not using a config file.")
        parser.add_argument("-v", "--verbose", type=int, help="Set verbosity level (0-2).", choices=[0, 1, 2])
        parser.add_argument("--log", type=str, help="Log file path.")

        args = parser.parse_args()

        if args.verbose == 0:
            logger.setLevel(logging.ERROR)
        elif args.verbose == 1:
            logger.setLevel(logging.INFO)
        elif args.verbose == 2:
            logger.setLevel(logging.DEBUG)
        else: # None
            logger.setLevel(logging.INFO)

        if args.log:
            addFileHandler(args.log)

        files = glob.glob(args.path, recursive=True)

        if len(files) == 0:
            raise FileNotFoundError(f"No files found matching {args.path}.")

        allJson = all([f.endswith(".json") for f in files])
        allXlsx = all([f.endswith(".xlsx") for f in files])

        if not (allJson or allXlsx):
            raise ValueError("All files must be either JSON or XLSX. Please use a glob pattern to select the correct files.")

        if allJson:
            configNames = files
            configs = []
            mode = "json"

            for configName in configNames:      
                try:
                    with open(configName, "r", encoding="utf-8") as f:
                        config = json.load(f)
                except json.JSONDecodeError as e:
                    raise ValueError(f"Config file {configName} is not a valid JSON file: {e}")
                except Exception as e:
                    raise ValueError(f"Error reading config file {configName}: {e}")
                
                config = cleanConfig(config)
                configs.append(config)

            logger.info(f"ExcelExtract version {version}.")
            logger.info(f"Found {len(configs)} config files.")

        else:
            mode = "xlsx"

            if not args.sheet:
                raise ValueError("When using XLSX files, you must specify a sheet name with -s/--sheet.")
            
            configs = [{
                    "input": files,
                    "simpletable": {
                        "sheet": args.sheet
                    }
                }]
            
            if not args.verbose and not args.output:
                logger.setLevel(logging.ERROR) # csv is printed to stdout, so set to silent

            logger.info(f"ExcelExtract version {version}.")

        for config in configs:
            if "exports" in config:
                exports = config["exports"]
            else:
                exports = [config]

            for exportConfig in exports:
                if mode == "json":
                    if args.input:
                        exportConfig["input"] = str(args.input)
                
                    if "output" not in exportConfig:
                        exportConfig["output"] = "output.csv"
                    if args.output:
                        exportConfig["output"] = args.output / exportConfig["output"]
                else:
                    if args.output:
                        exportConfig["output"] = args.output

                exportConfig = resolveSimpleTable(exportConfig)

                loopFiles(exportConfig)

        logger.info("Processing completed.")

    except FileNotFoundError as e:
        logger.error(f"{e}")
        sys.exit(1)

    except ValueError as e:
        logger.error(f"{e}")
        sys.exit(1)

    except Exception as e:
        logger.exception(f"{e}")
        sys.exit(1)

if __name__ == "__main__":
    main()
