
from openpyxl.utils import column_index_from_string, get_column_letter
from openpyxl.utils.cell import coordinate_from_string

from .tokens import applyTokenReplacement
from .lookup import resolveLookups
from .formulas import evaluate

def dereferenceCell(wb, cellRef, colSpec = None):
    if cellRef.strip().startswith("="):
        return evaluate(wb, cellRef)                
        
    else:
        # If the replaced value contains "!", treat it as a cell reference in the format "SheetName!CellRef".
        if "!" in cellRef:
            parts = cellRef.split("!", 1)
            refSheetName = parts[0]
            cellRef = parts[1]

            if colSpec is not None:
                if "rowoffset" in colSpec and colSpec["rowoffset"] != 0:
                    cellCoord = list(coordinate_from_string(cellRef))
                    cellCoord[1] += colSpec["rowoffset"]
                    cellRef = cellCoord[0] + str(cellCoord[1])
                if "coloffset" in colSpec and colSpec["coloffset"] != 0:
                    cellCoord = list(coordinate_from_string(cellRef))
                    cellCoord[0] = get_column_letter(column_index_from_string(cellCoord[0]) + colSpec["coloffset"])
                    cellRef = cellCoord[0] + str(cellCoord[1])

            try:
                sheet = wb[refSheetName]
                return sheet[cellRef].value
            except Exception as e:
                raise ValueError(f"Error reading cell {cellRef} from sheet {refSheetName}")
        else:
            return cellRef

def getColValue(wb, colDict, colName, tokens, cache, recursionDepth = 0):
    if recursionDepth > 100:
        raise ValueError("Recursion limit exceeded while resolving column value.")
    
    colSpec = colDict[colName]
    valueTemplate = colSpec.get("value", "")

    replacedValue = applyTokenReplacement(valueTemplate, tokens)

    # After the main tokens are replaced, check for other implicit column references.
    # Cache is only available for static columns, dynamic columns can not reference other dynamic columns.
    # That is, dynamic columns can only be referenced by static columns.
    if cache is not None:
        for cacheColName in cache:
            if "%%" + cacheColName + "%%" in replacedValue:
                cacheColValue = cache[cacheColName]
                if cacheColValue is None:
                    cacheColValue = 0
                replacedValue = replacedValue.replace("%%" + cacheColName + "%%", str(cacheColValue) if cacheColValue is not None else "")

    # If not in the cache, let's see if we can calculate it.
    #  - Static columns can reference other static columns.
    #  - Dynamic columns can reference other dynamic columns within the same intra-row token.
    #  - Dynamic columns can reference static columns (but the cache will remain none, thus they can not further reference other dynamic columns).
    colNames = colDict.keys()
    for otherColName in colNames:
        if otherColName != colName:
            if "%%" + otherColName + "%%" in replacedValue:
                otherColValue = getColValue(wb, colDict, otherColName, tokens, cache, recursionDepth + 1)

                otherType = colDict[otherColName].get("type", "string").lower()
                if otherType == "number":
                    if otherColValue is None:
                        otherColValue = 0

                replacedValue = replacedValue.replace("%%" + otherColName + "%%", str(otherColValue) if otherColValue is not None else "")

    return dereferenceCell(wb, replacedValue, colSpec)

def isDynamicCol(colNameDefinition, colValueDefinition, intraRowTokens):   
    for token in intraRowTokens:
        if "%%" + token + "%%" in colNameDefinition:
            return True
        if "%%" + token + "%%" in colValueDefinition:
            return True
    return False

def getColName(wb, colName, tokens):
    colName = applyTokenReplacement(colName, tokens)
    return dereferenceCell(wb, colName)

def checkForTrigger(colSpec, cellVal, colName):
    trigger = colSpec.get("trigger", "default").lower()
    if trigger not in ["default", "nonempty", "never", "nonzero"]:
        raise ValueError(f"Invalid trigger '{trigger}' for column '{colName}'. Valid triggers are 'default', 'nonempty', 'never', and 'nonzero'.")

    isEmpty = cellVal is None or cellVal == ""

    if trigger == "default" or trigger == "nonempty":
        if isEmpty:
            return False
        else:
            return True
        
    if trigger == "nonzero":
        if isEmpty:
            return False
        else:
            try:
                cellVal = float(cellVal)
        
                if cellVal is None:
                    return False
                elif cellVal != 0:
                    return True
                else:
                    return False
            except Exception:
                return False
        
    elif trigger == "never":
        return False
        
    raise RuntimeError(f"Unreachable code: trigger: {trigger}, colName: {colName}, cellVal: {cellVal}, type: {type(cellVal)}")

def extract(exportConfig, wb, filename):
    allRows = []
    types = {}

    if "columns" not in exportConfig:
        raise ValueError("Missing 'columns' in exportConfig")
    
    staticTokens =  {"FILE_NAME": filename}

    # Split the exportConfig into lookups and intra-row lookups.
    lookups = []
    intraRowLookups = []
    intraRowTokenNames = []
    if "lookups" in exportConfig:
        lookups = [lookup for lookup in exportConfig["lookups"] if "intrarow" not in lookup or lookup["intrarow"] == False]
        intraRowLookups = [lookup for lookup in exportConfig["lookups"] if "intrarow" in lookup and lookup["intrarow"] == True]
        intraRowTokenNames = [lookup["token"] for lookup in exportConfig["lookups"] if "intrarow" in lookup and lookup["intrarow"] == True]

    dynamicColumns = [
        colSpec["name"] 
        for colSpec in exportConfig["columns"] 
        if isDynamicCol(colSpec["name"], colSpec["value"], intraRowTokenNames)
    ]

    # Resolve the lookups to get the tokens for each row.
    tokensPerRow = []
    resolveLookups(wb, tokensPerRow, lookups, staticTokens)

    for tokens in tokensPerRow:
        rowData = {}
        triggerHit = False

        # Apply intra-row lookups to the current row's tokens.
        intraRowTokens = []
        resolveLookups(wb, intraRowTokens, intraRowLookups, tokens)

        # Each intraRowToken contains all static, row, and intra-row tokens for this row.
        # First do all dynamic columns, then static columns.
        for intraRowToken in intraRowTokens:    
            colDict = {getColName(wb, col["name"], intraRowToken): col for col in exportConfig["columns"]}

            for colName, colSpec in colDict.items():
                if colName is None or colName.strip() == "":
                    continue

                # Intra-row tokens can generate duplicate column names, just take the first one.
                if colName in rowData:
                    continue

                # We still need static columns in this list, but don't process them yet.
                if colSpec["name"] not in dynamicColumns:
                    continue

                cellVal = getColValue(wb, colDict, colName, intraRowToken, None)

                if checkForTrigger(colSpec, cellVal, colName):
                    triggerHit = True

                rowData[colName] = cellVal

                if "type" in colSpec:
                    if colName in types and types[colName] != colSpec["type"].lower():
                        raise ValueError(f"Column '{colName}' has conflicting types: '{types[colName]}' and '{colSpec['type'].lower()}'.")
                    else:
                        types[colName] = colSpec["type"].lower()
                else:
                    types[colName] = "auto"

        # Second pass for static columns.  
        colDictStatic = {
            getColName(wb, col["name"], tokens): col 
            for col in exportConfig["columns"] if col["name"] not in dynamicColumns
        }

        for colName, colSpec in colDictStatic.items():
            if colName is None or colName.strip() == "":
                continue
            
            # Intra-row tokens can generate duplicate column names, just take the first one.
            if colName in rowData:
                continue

            cellVal = getColValue(wb, colDictStatic, colName, tokens, rowData)

            if checkForTrigger(colSpec, cellVal, colName):
                triggerHit = True

            rowData[colName] = cellVal

            if "type" in colSpec:
                if colName in types and types[colName] != colSpec["type"].lower():
                    raise ValueError(f"Column '{colName}' has conflicting types: '{types[colName]}' and '{colSpec['type'].lower()}'.")
                else:
                    types[colName] = colSpec["type"].lower()
            else:
                types[colName] = "auto"

        # Only add the row if at least one cell hits the trigger condition.
        if triggerHit:
            allRows.append(rowData)

    return allRows, types
