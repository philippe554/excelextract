import subprocess
from openpyxl import Workbook
import json

def generateTestData(path):
    wb = Workbook()
    ws = wb.active

    ws.title = "sheet1"
    ws.append(["Name", "Age", "City"])
    ws.append(["Alice", 30, "New York"])
    ws.append(["Bob", 25, "Los Angeles"])
    ws.append(["Charlie", 35, "Chicago"])

    wb.save(path / "test_data.xlsx")

    config = {
        "exports": [
            {
                "input": str(path / "test_data.xlsx"),
                "output": str(path / "output.csv"),
                "lookups": [
                    {
                        "operation": "looprows",
                        "start": 2,
                        "end": 10,
                        "token": "ROW",
                    }
                ],
                "columns": [
                    {"name": "File", "type": "string", "value": "%%FILE_NAME%%", "trigger": "never"},
                    {"name": "Name", "type": "string", "value": "sheet1!A%%ROW%%"},
                    {"name": "Age", "type": "number", "value": "sheet1!B%%ROW%%"},
                    {"name": "City", "type": "string", "value": "sheet1!C%%ROW%%"},
                ]
            }
        ]
    }

    with open(path / "config.json", "w") as f:
        json.dump(config, f, indent=4)

def test_cli_help():
    result = subprocess.run(["excelextract", "--help"], capture_output=True, text=True)

    print("STDOUT:", result.stdout)
    print("STDERR:", result.stderr)

    assert result.returncode == 0

def test_cli_config(tmp_path):
    generateTestData(tmp_path)

    result = subprocess.run(["excelextract", str(tmp_path / "config.json")], capture_output=True, text=True)

    print("STDOUT:", result.stdout)
    print("STDERR:", result.stderr)

    with open(tmp_path / "output.csv", "r") as f:
        output = f.read()
    print("Output CSV:", output)

    assert result.returncode == 0
    assert len(output.splitlines()) == 4