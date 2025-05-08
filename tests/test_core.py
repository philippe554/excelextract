import subprocess

def test_1(tmp_path):
    result = subprocess.run(["excelextract", "tests/data/config.json", "-i", "tests/data/*.xlsx", "-o", tmp_path], capture_output=True, text=True)
    print("STDOUT:\n" + result.stdout)
    print("STDERR:\n" + result.stderr)
    assert result.returncode == 0

    with open(tmp_path / "employees.csv", "r") as f:
        output = f.read()
    print("Employees CSV:\n" + output)
    assert len(output.splitlines()) == 10

    with open(tmp_path / "inventory.csv", "r") as f:
        output = f.read()
    print("Inventory CSV:\n" + output)
    assert len(output.splitlines()) == 17

    with open(tmp_path / "findcell.csv", "r") as f:
        output = f.read()
    print("Find Cell CSV:\n" + output)
    assert output.splitlines()[1] == "\"12\",\"F\""

    with open(tmp_path / "formulas.csv", "r") as f:
        output = f.read()
    print("Formulas CSV:\n" + output)
    assert output.splitlines()[1] == "33.0,1.5,\"Hello World\",\"45200\",\"True\",\"Yes\""

    with open(tmp_path / "implicit.csv", "r") as f:
        output = f.read()
    print("Implicit CSV:\n" + output)
    assert output.splitlines()[1] == "\"surveys.xlsx\",4.0,40.0,160.0"

    with open(tmp_path / "basic.csv", "r") as f:
        output = f.read()
    print("Basic CSV:\n" + output)
    assert output.splitlines()[1] == "\"Alice\",\"Engineer\",\"30000\",\"4\",\"USD\",\"40\""
    assert len(output.splitlines()) == 8

    with open(tmp_path / "simpleTable.csv", "r") as f:
        output2 = f.read()
    print("Simple Table CSV:\n" + output2)
    assert output == output2

