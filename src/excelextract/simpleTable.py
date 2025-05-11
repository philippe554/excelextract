
from openpyxl.utils import column_index_from_string, get_column_letter

def resolveSimpleTable(config):
    try:
        if "simpletable" not in config:
            return config
        
        if "sheet" not in config["simpletable"]:
            raise ValueError("Missing 'sheet' in simpletable configuration")
        
        sheet = config["simpletable"]["sheet"]
        headerRow = config["simpletable"].get("headerrow", 1)
        startRow = config["simpletable"].get("startrow", headerRow + 1) 
        endRow = config["simpletable"].get("endrow", None)
        if endRow is not None:
            endRow -= 1
        count = config["simpletable"].get("count", None)
        if count is not None:
            endRow = startRow + count - 1
        startCol = config["simpletable"].get("startcol", "A")
        endCol = config["simpletable"].get("endcol", None)
        sourceFileColumn = config["simpletable"].get("sourcefilecolumn", None)
        sheetNameColumn = config["simpletable"].get("sheetnamecolumn", None)

        del config["simpletable"]
        
        if "lookups" not in config:
            config["lookups"] = []

        config["lookups"].append(
            {
                "operation": "looprows",
                "start": startRow,
                "end": endRow,
                "token": "__SIMPLETABLE_ROW"
            })
        config["lookups"].append(
            {
                "operation": "loopcolumns",
                "start": startCol,
                "end": endCol,
                "token": "__SIMPLETABLE_COL",
                "intrarow": True
            })
        
        if "columns" not in config:
            config["columns"] = []

        if sourceFileColumn is not None:
            config["columns"].append(
                {
                    "name": sourceFileColumn,
                    "type": "string",
                    "value": "%%FILE_NAME%%",
                    "trigger": "never"
                })
            
        if sheetNameColumn is not None:
            config["columns"].append(
                {
                    "name": sheetNameColumn,
                    "type": "string",
                    "value": sheet,
                    "trigger": "never"
                })
            
        config["columns"].append(
            {
                "name": sheet + "!%%__SIMPLETABLE_COL%%" + str(headerRow),
                "type": "auto",
                "value": sheet + "!%%__SIMPLETABLE_COL%%%%__SIMPLETABLE_ROW%%"
            })
        
        print(config)
        
        return config
    
    except Exception as e:
        raise ValueError(f"Error resolving simple table configuration: {e}")