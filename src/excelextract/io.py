
import os
import csv
import glob
import warnings

from openpyxl import load_workbook

from .extract import extract
from .type import detectTypeOfList, convertRowToType

def loopFiles(exportConfig):
    if "input" not in exportConfig:
        raise ValueError("Missing 'input' in exportConfig")
    inputGlobs = exportConfig["input"]
    if not isinstance(inputGlobs, list):
        inputGlobs = [inputGlobs]

    inputFiles = []

    for inputGlob in inputGlobs:
        if not isinstance(inputGlob, str):
            raise ValueError(f"Invalid input glob: {inputGlob}")
        inputFiles.extend(glob.glob(inputGlob, recursive=True))

    if not inputFiles:
        raise ValueError(f"No files found matching input glob(s): {inputGlobs}")       
        
    allRows = []
    allTypes = {}

    for inputFile in inputFiles:
        if not os.path.isfile(inputFile):
            raise ValueError(f"Input file not found: {inputFile}")
        if not inputFile.endswith(".xlsx"):
            raise ValueError(f"Input file is not an Excel file: {inputFile}")
        inputFileName = os.path.basename(inputFile)
        if inputFileName.startswith("~$"):
            continue
        
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore", category=UserWarning)
                wb = load_workbook(inputFile, data_only=True)
        except Exception as e:
            raise ValueError(f"Error opening file {inputFile}: {e}")
        
        rows, types = extract(exportConfig, wb, inputFileName)

        print(f"Processing {inputFile} with {len(rows)} rows extracted.")

        allRows.extend(rows)

        for key, value in types.items():
            if key not in allTypes:
                allTypes[key] = value
            else:
                if allTypes[key] != value:
                    raise ValueError(f"Column type mismatch for column '{key}' in file '{inputFile}': {allTypes[key]} vs {value}")

    if len(allRows) == 0:
        raise ValueError("No rows extracted from the input files")
    
    if "output" not in exportConfig:
        raise ValueError("Missing 'output' in exportConfig")
    
    outputFile = exportConfig["output"]

    if not str(outputFile).endswith(".csv"):
        outputFile += ".csv"

    outputDir = os.path.dirname(outputFile)
    if outputDir != "" and not os.path.exists(outputDir):
        os.makedirs(outputDir)

    with open(outputFile, "w", newline="", encoding="utf-8-sig") as csvfile:       
        colNames = allTypes.keys()
        
        for colName in colNames:
            if allTypes[colName] == "auto":
                data = [d[colName] for d in allRows if colName in d]
                allTypes[colName] = detectTypeOfList(data)
        
        writer = csv.DictWriter(csvfile, fieldnames = colNames, delimiter=",", quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        writer.writeheader()
        for row in allRows:
            rowConverted = convertRowToType(row, allTypes)
            writer.writerow(rowConverted)

    print(f"Wrote {len(allRows)} rows to {outputFile}")
        
