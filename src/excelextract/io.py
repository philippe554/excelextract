
import os
import csv
import glob

from openpyxl import load_workbook

from .extract import extract

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

    for inputFile in inputFiles:
        if not os.path.isfile(inputFile):
            raise ValueError(f"Input file not found: {inputFile}")
        if not inputFile.endswith(".xlsx"):
            raise ValueError(f"Input file is not an Excel file: {inputFile}")
        inputFileName = os.path.basename(inputFile)
        if inputFileName.startswith("~$"):
            continue

        print(f"Processing file: {inputFile}")
        
        try:
            wb = load_workbook(inputFile, data_only=True)
        except Exception as e:
            raise ValueError(f"Error opening file {inputFile}: {e}")
        
        rows = extract(exportConfig, wb, inputFileName)
        allRows.extend(rows)

    if len(allRows) == 0:
        raise ValueError("No rows extracted from the input files")
    
    if "output" not in exportConfig:
        raise ValueError("Missing 'output' in exportConfig")
    
    outputFile = exportConfig["output"]

    if not str(outputFile).endswith(".csv"):
        outputFile += ".csv"

    outputDir = os.path.dirname(outputFile)
    if not os.path.exists(outputDir):
        os.makedirs(outputDir)

    with open(outputFile, "w", newline="", encoding="utf-8-sig") as csvfile:
        fieldNames = [col["name"] for col in exportConfig.get("columns", [])]
        writer = csv.DictWriter(csvfile, fieldnames = fieldNames, delimiter=",", quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
        writer.writeheader()
        for row in allRows:
            writer.writerow(row)
        
