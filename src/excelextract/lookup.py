
import re

from openpyxl.utils import column_index_from_string, get_column_letter

from .tokens import applyTokenReplacement

def resolveLookups(wb, elements = [], unprocessedDefinitions = [], currentElement = {}):
    if len(unprocessedDefinitions) == 0:
        elements.append(currentElement)
    else:
        loopDefinition = unprocessedDefinitions[0]

        type = loopDefinition.get("type", "unknown").lower()
        if type not in ["loopsheets", "findrow", "findcol", "looprowsfixed", "loopcolsfixed", "looprowsbysearch", "loopcolsbysearch"]:
            raise ValueError(f"Invalid loop type '{type}' in definition: {loopDefinition}")
        
        token = loopDefinition["token"]
        loopElements = []
        
        if type == "loopsheets":
            sheetRegex = loopDefinition.get("regex", ".*")
            matchingSheets = [sheet.title for sheet in wb.worksheets if re.search(sheetRegex, sheet.title)]
            loopElements = matchingSheets

        elif type == "findrow":
            regex = loopDefinition.get("regex", ".*")
            sheet = applyTokenReplacement(loopDefinition["sheet"], currentElement)
            column = applyTokenReplacement(loopDefinition.get("column", "A"), currentElement)
            offset = loopDefinition.get("offset", 0)
            data = [str(cell.value) if cell.value is not None else "" for cell in wb[sheet][column]]
            index = next((i for i, s in enumerate(data) if re.search(regex, s)), -1)
            if index == -1:
                raise ValueError(f"Search regex '{regex}' not found in column '{column}' of sheet '{sheet}'")
            index += offset + 1 # +1 to convert to 1-based index
            loopElements = [index]

        elif type == "findcol":
            regex = loopDefinition.get("regex", ".*")
            sheet = applyTokenReplacement(loopDefinition["sheet"], currentElement)
            row = applyTokenReplacement(loopDefinition.get("row", 1), currentElement)
            offset = loopDefinition.get("offset", 0)
            data = [str(cell.value) if cell.value is not None else "" for cell in wb[sheet][row]]
            index = next((i for i, s in enumerate(data) if re.search(regex, s)), -1)
            if index == -1:
                raise ValueError(f"Search regex '{regex}' not found in row '{row}' of sheet '{sheet}'")
            index += offset + 1
            loopElements = [get_column_letter(index)]

        elif type == "looprowsfixed":
            startRow = int(loopDefinition.get("start", 1))
            endRow = int(loopDefinition.get("end", 1))
            stride = loopDefinition.get("stride", 1)
            rowIndices = list(range(startRow, endRow + 1, stride))
            loopElements = rowIndices

        elif type == "loopcolsfixed":
            startCol = column_index_from_string(loopDefinition.get("start", "A"))
            endCol = column_index_from_string(loopDefinition.get("end", "A"))
            stride = loopDefinition.get("stride", 1)
            columnIndices = list(range(startCol, endCol + 1, stride))
            loopElements = [get_column_letter(i) for i in columnIndices]

        elif type == "looprowsbysearch":
            searchSheet = applyTokenReplacement(loopDefinition["sheet"], currentElement)

            startRegex = loopDefinition["start"].get("regex", ".*")
            startCol = applyTokenReplacement(loopDefinition["start"].get("column", "A"), currentElement)
            startOffset = loopDefinition["start"].get("offset", 0)
            startData = [str(cell.value) if cell.value is not None else "" for cell in wb[searchSheet][startCol]]
            startIndex = next((i for i, s in enumerate(startData) if re.search(startRegex, s)), -1)
            if startIndex == -1:
                raise ValueError(f"Start search regex '{startRegex}' not found in column '{startCol}' of sheet '{searchSheet}'")
            startIndex += startOffset + 1 # +1 to convert to 1-based index

            if "regex" in loopDefinition["end"]:
                endRegex = loopDefinition["end"].get("regex", ".*")
                endCol = applyTokenReplacement(loopDefinition["end"].get("column", "A"), currentElement)
                endOffset = loopDefinition["end"].get("offset", 0)
                endData = [str(cell.value) if cell.value is not None else "" for cell in wb[searchSheet][endCol]]
                endIndex = next((i for i, s in enumerate(endData) if re.search(endRegex, s)), -1)
                if endIndex == -1:
                    raise ValueError(f"End search regex '{endRegex}' not found in column '{endCol}' of sheet '{searchSheet}'")
                endIndex += endOffset + 1 # +1 to convert to 1-based index

            if "count" in loopDefinition["end"]:
                endCount = loopDefinition["end"].get("count", 0)
                endIndex = startIndex + endCount - 1

            if startIndex > endIndex:
                raise ValueError(f"Start index {startIndex} is greater than end index {endIndex} for sheet '{searchSheet}'")
            
            stride = loopDefinition.get("stride", 1)
            loopElements = list(range(startIndex, endIndex + 1, stride))

        elif type == "loopcolsbysearch":
            searchSheet = applyTokenReplacement(loopDefinition["sheet"], currentElement)

            startRegex = loopDefinition["start"].get("regex", ".*")
            startRow = applyTokenReplacement(loopDefinition["start"].get("row", 1), currentElement)
            startOffset = loopDefinition["start"].get("offset", 0)
            startData = [str(cell.value) if cell.value is not None else "" for cell in wb[searchSheet][startRow]]
            startIndex = next((i for i, s in enumerate(startData) if re.search(startRegex, s)), -1)
            if startIndex == -1:
                raise ValueError(f"Start search regex '{startRegex}' not found in row '{startRow}' of sheet '{searchSheet}'")
            startIndex += startOffset + 1 # +1 to convert to 1-based index

            if "regex" in loopDefinition["end"]:
                endRegex = loopDefinition["end"].get("regex", ".*")
                endRow = applyTokenReplacement(loopDefinition["end"].get("row", 1), currentElement)
                endOffset = loopDefinition["end"].get("offset", 0)
                endData = [str(cell.value) if cell.value is not None else "" for cell in wb[searchSheet][endRow]]
                endIndex = next((i for i, s in enumerate(endData) if re.search(endRegex, s)), -1)
                if endIndex == -1:
                    raise ValueError(f"End search regex '{endRegex}' not found in row '{endRow}' of sheet '{searchSheet}'")
                endIndex += endOffset + 1 # +1 to convert to 1-based index

            if "count" in loopDefinition["end"]:
                endCount = loopDefinition["end"].get("count", 0)
                endIndex = startIndex + endCount - 1

            if startIndex > endIndex:
                raise ValueError(f"Start index {startIndex} is greater than end index {endIndex} for sheet '{searchSheet}'")
            stride = loopDefinition.get("stride", 1)

            loopElements = [get_column_letter(i) for i in range(startIndex, endIndex + 1, stride)]

        if len(loopDefinition) == 0:
            raise ValueError("Loop definition is empty")

        for i in range(len(loopElements)):
            copy = currentElement.copy()
            copy[token] = loopElements[i]
            resolveLookups(wb, elements, unprocessedDefinitions[1:], copy)
