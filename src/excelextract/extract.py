
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from .tokens import applyTokenReplacement
from .lookup import resolveLookups

def extract(exportConfig, wb, filename):
    allRows = []

    if "columns" not in exportConfig:
        raise ValueError("Missing 'columns' in exportConfig")

    tokensPerRow = []
    if "lookups" in exportConfig:
        resolveLookups(wb, tokensPerRow, exportConfig["lookups"], {"FILE_NAME": filename})
    if len(tokensPerRow) == 0:
        tokensPerRow = [{"FILE_NAME": filename}]

    for tokens in tokensPerRow:
        rowData = {}
        triggerHit = False        

        # For each column in the configuration, perform token replacement.
        for col in exportConfig["columns"]:
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

                if "rowoffset" in col and col["rowoffset"] != 0:
                    cellCoord = list(coordinate_from_string(cellRef))
                    cellCoord[1] += col["rowoffset"]
                    cellRef = cellCoord[0] + str(cellCoord[1])
                if "coloffset" in col and col["coloffset"] != 0:
                    cellCoord = list(coordinate_from_string(cellRef))
                    cellCoord[0] += get_column_letter(column_index_from_string(cellCoord[0]) + col["coloffset"])
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

    return allRows
