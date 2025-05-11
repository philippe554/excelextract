
import datetime
from dateutil.parser import parse

ALL_SUPPORTED_TYPES = {"string", "number", "integer", "float", "boolean", "date", "datetime", "time", "timedelta"}

# returns a tuple of (parsed object, type)
def analyseDateTimeString(value):
    default1 = datetime.datetime(1, 1, 1, 0, 0, 0, 0)
    default2 = datetime.datetime(2, 2, 2, 2, 2, 2, 2)

    try:
        dt1 = parse(value, default=default1)
        dt2 = parse(value, default=default2)

        hasDate = True
        if dt1.year == 1 and dt1.month == 1 and dt1.day == 1 and dt2.year == 2 and dt2.month == 2 and dt2.day == 2:
            hasDate = False

        hasTime = True
        if dt1.hour == 0 and dt1.minute == 0 and dt1.second == 0 and dt2.hour == 2 and dt2.minute == 2 and dt2.second == 2:
            hasTime = False

        if hasDate and hasTime:
            if dt1.hour == 0 and dt1.minute == 0 and dt1.second == 0 and dt1.microsecond == 0:
                return dt1, "date"
            else:
                return dt1, "datetime"
        elif hasDate:
            return dt1, "date"
        elif hasTime:
            return dt1, "time"
        else:
            return None, None

    except ValueError:
        return None, None

def detectAllPossibleTypes(value):
    if value is None:
        return ALL_SUPPORTED_TYPES
    
    # Handle all cases where we know the type

    if isinstance(value, bool):
        return {"string", "number", "float", "integer", "boolean"}
    
    if isinstance(value, int):
        if value in [0, 1]:
            return {"string", "number", "float", "integer", "boolean"}
        else:
            return {"string", "number", "float", "integer"}
    
    if isinstance(value, float):
        if value.is_integer():
            return detectAllPossibleTypes(int(value))
        else:
            return {"string", "number", "float"}
    
    if isinstance(value, datetime.datetime):
        if value.hour == 0 and value.minute == 0 and value.second == 0 and value.microsecond == 0:
            return {"string", "date", "datetime"}
        else:
            return {"string", "datetime"} # datetime can't be converted to date without losing information
    
    if isinstance(value, datetime.date):
        return {"string", "date", "datetime"}
        
    if isinstance(value, datetime.time):
        return {"string", "time"} # time can't be converted to datetime without extra information
    
    if isinstance(value, datetime.timedelta):
        return {"string", "timedelta"}
        
    # We have an unknown type that is not a string
    # As long as we can convert it to a string, ExcelExtract will accept it as a string
    if not isinstance(value, str):
        try:
            value = str(value)
            return {"string"}
        except Exception:
            raise ValueError(f"Unknown type and cannot be converted to string: {value} ({type(value)})")
    
    # From here on, we have a string

    if value == "":
        return ALL_SUPPORTED_TYPES
    
    if value.lower() in ("true", "t", "false", "f"):
        return {"string", "boolean"}
    
    obj, type = analyseDateTimeString(value)
    if obj is not None:
        if type == "date":
            return {"string", "date", "datetime"}
        elif type == "datetime":
            return {"string", "datetime"} # datetime can't be converted to date without losing information
        elif type == "time":
            return {"string", "time"} # time can't be converted to datetime without extra information
        
        # timedelta can not be automatically detected

    try:
        return detectAllPossibleTypes(float(value))
    except (ValueError, TypeError):
        pass

    return {"string"}

def intersectionOfSets(sets):
    if len(sets) == 0:
        return set()
    if len(sets) == 1:
        return sets[0]
    else:
        result = sets[0]

        for s in sets[1:]:
            result = result.intersection(s)

        return result

def getMostSpecificType(set):
    if "boolean" in set:
        return "boolean"
    
    if "date" in set: # if everything can be represented as date, we prefer date over datetime
        return "date"
    if "time" in set:
        return "time"
    if "timedelta" in set:
        return "timedelta"
    if "datetime" in set:
        return "datetime"

    if "integer" in set:
        return "integer"
    if "float" in set:
        return "float"
    if "number" in set:
        return "number"
    
    if "string" in set:
        return "string"
    
    raise ValueError(f"Unknown type set: {set}")

def detectTypeOfList(data):
    possibleTypesPerValue = [detectAllPossibleTypes(value) for value in data]

    possibleTypes = intersectionOfSets(possibleTypesPerValue)

    return getMostSpecificType(possibleTypes)


def convertRowToType(row: dict, colTypes: dict) -> dict:
    for key, value in row.items():
        if key in colTypes:
            if value is None or value == "":
                row[key] = None

            elif colTypes[key] == "string":
                row[key] = str(value)

            elif colTypes[key] == "boolean":
                if isinstance(value, bool):
                    row[key] = value
                elif isinstance(value, str):
                    if value.lower() in ("true", "t"):
                        row[key] = True
                    elif value.lower() in ("false", "f"):
                        row[key] = False
                    else:
                        row[key] = None
                else:
                    try:
                        row[key] = bool(int(value))
                    except ValueError:
                        row[key] = None

            elif colTypes[key] == "number" or colTypes[key] == "float":
                try:
                    row[key] = float(value)
                except ValueError:
                    row[key] = None

            elif colTypes[key] == "integer":
                try:
                    row[key] = int(value)
                except ValueError:
                    row[key] = None

            elif colTypes[key] == "datetime":
                if isinstance(value, datetime.datetime):
                    row[key] = str(value)
                elif isinstance(value, datetime.date): # assume time is 00:00:00
                    row[key] = str(datetime.datetime.combine(value, datetime.datetime.min.time()))
                elif isinstance(value, str):
                    obj, type = analyseDateTimeString(value)
                    if obj is None:
                        row[key] = None
                    else:
                        row[key] = str(obj)
                else:
                    row[key] = None

            elif colTypes[key] == "date":
                if isinstance(value, datetime.datetime): # Only relevant if user specified date type
                    row[key] = str(value.date())
                elif isinstance(value, datetime.date):
                    row[key] = str(value)
                elif isinstance(value, str):
                    obj, type = analyseDateTimeString(value)
                    if obj is None:
                        row[key] = None
                    else:
                        row[key] = str(obj.date())
                else:
                    row[key] = None

            elif colTypes[key] == "time":
                if isinstance(value, datetime.datetime): # Only relevant if user specified time type
                    row[key] = str(value.time())
                elif isinstance(value, datetime.time):
                    row[key] = str(value)
                elif isinstance(value, str):
                    obj, type = analyseDateTimeString(value)
                    if obj is None:
                        row[key] = None
                    else:
                        row[key] = str(obj.time())
                else:
                    row[key] = None

            elif colTypes[key] == "timedelta":
                if isinstance(value, datetime.timedelta):
                    row[key] = str(value)
                elif isinstance(value, str):
                    row[key] = value
                else:
                    row[key] = None

            else:
                raise ValueError(f"Unknown type {colTypes[key]} for column {key} (target type: {colTypes[key]})")

    return row
