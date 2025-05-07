
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from .tokens import applyTokenReplacement
from .lookup import resolveLookups
from .formulas import evaluate

def getColValue(wb, colDict, colName, tokens, recursionDepth = 0):
    if recursionDepth > 100:
        raise ValueError("Recursion limit exceeded while resolving column value.")
    
    colSpec = colDict[colName]
    valueTemplate = colSpec.get("value", "")

    replacedValue = applyTokenReplacement(valueTemplate, tokens)

    # After the main tokens are replaced, check for other column references.
    colNames = colDict.keys()
    for otherColName in colNames:
        if otherColName != colName:
            if "%%" + otherColName + "%%" in replacedValue:
                otherColValue = getColValue(wb, colDict, otherColName, tokens, recursionDepth + 1)

                otherType = colDict[otherColName].get("type", "string").lower()
                if otherType == "number":
                    if otherColValue is None:
                        otherColValue = 0

                replacedValue = replacedValue.replace("%%" + otherColName + "%%", str(otherColValue) if otherColValue is not None else "")

    if replacedValue.strip().startswith("="):
        return evaluate(wb, replacedValue)                
        
    else:
        # If the replaced value contains "!", treat it as a cell reference in the format "SheetName!CellRef".
        if "!" in replacedValue:
            parts = replacedValue.split("!", 1)
            refSheetName = parts[0]
            cellRef = parts[1]

            if "rowoffset" in colSpec and colSpec["rowoffset"] != 0:
                cellCoord = list(coordinate_from_string(cellRef))
                cellCoord[1] += colSpec["rowoffset"]
                cellRef = cellCoord[0] + str(cellCoord[1])
            if "coloffset" in colSpec and colSpec["coloffset"] != 0:
                cellCoord = list(coordinate_from_string(cellRef))
                cellCoord[0] += get_column_letter(column_index_from_string(cellCoord[0]) + colSpec["coloffset"])
                cellRef = cellCoord[0] + str(cellCoord[1])

            try:
                sheet = wb[refSheetName]
                return sheet[cellRef].value
            except Exception as e:
                raise ValueError("Error reading cell {cellRef} from sheet {refSheetName}")
        else:
            return replacedValue

def extract(exportConfig, wb, filename):
    allRows = []

    if "columns" not in exportConfig:
        raise ValueError("Missing 'columns' in exportConfig")

    tokensPerRow = []
    if "lookups" in exportConfig:
        resolveLookups(wb, tokensPerRow, exportConfig["lookups"], {"FILE_NAME": filename})
    if len(tokensPerRow) == 0:
        tokensPerRow = [{"FILE_NAME": filename}]

    colDict = {col["name"]: col for col in exportConfig["columns"]}

    for tokens in tokensPerRow:
        rowData = {}
        triggerHit = False        

        for colName, colSpec in colDict.items():
            cellVal = getColValue(wb, colDict, colName, tokens)

            trigger = colSpec.get("trigger", "default").lower()
            if trigger not in ["default", "nonempty", "never", "nonzero"]:
                print(f"  Error: Invalid trigger '{trigger}' for column '{colName}'. Using default trigger.")
                trigger = "default"

            colType = colSpec.get("type", "string").lower()
            if colType not in ["string", "number"]:
                print(f"  Error: Invalid type '{colType}' for column '{colName}'. Using default type 'string'.")
                colType = "string"

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

    return allRows
