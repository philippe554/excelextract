
import os
import re
import csv

from openpyxl import load_workbook
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from .tokens import applyTokenReplacement
from .lookup import resolveLookups

def extract(exportConfig):
    inputFolder = exportConfig.get("inputFolder", ".")
    inputRegex = exportConfig.get("inputRegex", ".*")
    outputName = exportConfig.get("output", "output")

    print(f"Processing export: {outputName} ...")

    if not outputName.endswith(".csv"):
        outputName += ".csv"

    # Find Excel files matching the input regex in the given folder.
    files = [f for f in os.listdir(inputFolder) if re.search(inputRegex, f)]
    if not files:
        print(f"  Error: No files found in {inputFolder} matching pattern {inputRegex}")
        return

    allRows = []  # This will hold all rows for output

    for filename in files:
        if filename.startswith("~$"):
            continue
        if not filename.endswith(".xlsx"):
            continue

        filepath = os.path.join(inputFolder, filename)
        print(f"  Info: Processing file: {filepath}")
        try:
            wb = load_workbook(filepath, data_only=True)

        except Exception as e:
            print(f"  Error: Error opening file {filepath}: {e}")
            continue

        tokensPerRow = []
        resolveLookups(wb, tokensPerRow, exportConfig.get("lookups", []), {"FILE_NAME": filename})

        if len(tokensPerRow) == 0:
            tokensPerRow = [{"FILE_NAME": filename}]

        for tokens in tokensPerRow:
            rowData = {}
            triggerHit = False        

            # For each column in the configuration, perform token replacement.
            for col in exportConfig.get("columns", []):
                colName = col.get("name")

                colType = col.get("type", "string").lower()
                if colType not in ["string", "number"]:
                    print(f"  Error: Invalid type '{colType}' for column '{colName}'. Using default type 'string'.")
                    colType = "string"

                valueTemplate = col.get("value", "")

                trigger = col.get("trigger", "default").lower()
                if trigger not in ["default", "nonempty", "never", "nonzero"]:
                    print(f"  Error: Invalid trigger '{trigger}' for column '{colName}'. Using default trigger.")
                    trigger = "default"

                replacedValue = applyTokenReplacement(valueTemplate, tokens)

                # If the replaced value contains "!", treat it as a cell reference in the format "SheetName!CellRef".
                if "!" in replacedValue:
                    parts = replacedValue.split("!", 1)
                    refSheetName = parts[0]
                    cellRef = parts[1]

                    if "rowOffset" in col and col["rowOffset"] != 0:
                        cellCoord = list(coordinate_from_string(cellRef))
                        cellCoord[1] += col["rowOffset"]
                        cellRef = cellCoord[0] + str(cellCoord[1])
                    if "colOffset" in col and col["colOffset"] != 0:
                        cellCoord = list(coordinate_from_string(cellRef))
                        cellCoord[0] += get_column_letter(column_index_from_string(cellCoord[0]) + col["colOffset"])
                        cellRef = cellCoord[0] + str(cellCoord[1])

                    try:
                        sheet = wb[refSheetName]
                        cellVal = sheet[cellRef].value
                    except Exception as e:
                        print(f"  Error: Error reading cell {cellRef} from sheet {refSheetName} in file {filename}: {e}")
                        cellVal = None
                else:
                    cellVal = replacedValue

                # Convert the cell value according to the specified type.
                if colType == "number":
                    try:
                        isEmpty = cellVal is None or cellVal == ""
                            
                        cellVal = float(cellVal) if not isEmpty else None

                        if trigger == "nonzero" and not isEmpty:
                            if cellVal != 0:
                                triggerHit = True

                    except Exception:
                        cellVal = None
                elif colType == "string":
                    if trigger == "nonzero":
                        print(f"  Error: Nonzero trigger is not applicable for string type in column '{colName}'. Using default trigger.")

                    if cellVal is not None:
                        cellVal = str(cellVal)

                rowData[colName] = cellVal

                if trigger == "default" or trigger == "nonempty":
                    if cellVal not in [None, ""]:
                        triggerHit = True

            # Only add the row if at least one cell hits the trigger condition.
            if triggerHit:
                allRows.append(rowData)

    # Write all extracted rows to CSV.
    if allRows:
        print(f"  Info: Writing output to {outputName} ...")
        with open(outputName, "w", newline="", encoding="utf-8-sig") as csvfile:
            fieldNames = [col["name"] for col in exportConfig.get("columns", [])]
            writer = csv.DictWriter(csvfile, fieldnames = fieldNames, delimiter=",", quotechar='"', quoting=csv.QUOTE_NONNUMERIC)
            writer.writeheader()
            for row in allRows:
                writer.writerow(row)
    else:
        print("  Warning: No rows to write.")
    print()
