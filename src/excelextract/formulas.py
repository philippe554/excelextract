
import formulas
from openpyxl.utils.cell import rows_from_range

def getCellValue(wb, cellRef, default=None):
    if "!" not in cellRef:
        raise ValueError("Cell reference must be in the format 'SheetName!CellRef': " + cellRef)
    
    parts = cellRef.split("!", 1)

    refSheetName = parts[0].strip("'")
    cellRef = parts[1]
    
    sheetNames = [name.upper() for name in wb.sheetnames]
    if refSheetName.upper() not in sheetNames:
        raise ValueError("Sheet name not found: " + refSheetName)
    sheetIndex = sheetNames.index(refSheetName.upper())
    caseSensitiveSheetName = wb.sheetnames[sheetIndex]

    value = wb[caseSensitiveSheetName][cellRef].value

    if value is None:
        return default
    else:
        return value

def evaluate(wb, formula):
    formula = formula.strip()
    if not formula.startswith('='):
        raise ValueError("Formula must start with '='")
    
    try:
        func = formulas.Parser().ast(formula)[1].compile()

        inputsRefs = list(func.inputs)

        inputs = []

        for ref in inputsRefs:
            if "!" not in ref:
                raise ValueError("Cell reference must be in the format 'SheetName!CellRef': " + ref)
            
            parts = ref.split("!", 1)
            refSheetName = parts[0]
            cellRef = parts[1]

            if ":" in cellRef:
                rangeElementRef = rows_from_range(cellRef)
                data = []
                for rangeElement in rangeElementRef:
                    value = getCellValue(wb, refSheetName + "!" + rangeElement[0])
                    data.append(value)
                inputs.append(data)
            else:
                inputs.append(getCellValue(wb, refSheetName + "!" + cellRef))

        return func(*inputs)

    except Exception as e:
        raise ValueError(f"Error parsing formula: {formula}. Error: {str(e)}")

