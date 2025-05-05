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

