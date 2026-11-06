"""Record independent scalar formula results from pinned Calc, without the SDK."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("scalar_native_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-scalar-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-scalar-semantics.xlsx"
FORMULAS = [
    '1+2*3', '-2^2', '2^3^2', '5%', '"2"+3', '""+1', 'TRUE+2', '"A"&TRUE', '1/0', 'SQRT(-1)',
    '1="1"', 'TRUE=1', '2<"1"', '"a"="A"', 'A5=0', 'A5=""', 'A5=FALSE', 'A4=0', 'A4=FALSE', 'A4=A5',
    'A5', 'C1', 'C2', 'IF(TRUE,1,1/0)', 'IF(FALSE,1/0,2)', 'IF(1,,)', 'IF(FALSE,1)', 'IF(TRUE,A5,1)',
    'IFERROR(1/0,3)', '_xlfn.IFNA(1/0,3)', '_xlfn.IFNA(A6,3)', 'IF("x",1,2)', 'IF("TRUE",1,2)',
    'AND(A1:A5)', 'OR(A1:A5)', 'AND(A2:A2)', 'AND("x",TRUE)', 'AND(TRUE,"TRUE")', 'NOT("TRUE")', 'NOT(0)',
    'AND(FALSE,1/0)', 'OR(TRUE,1/0)', 'TRUE()', 'FALSE()', 'ISBLANK(A5)', 'ISBLANK(A4)', 'ISNUMBER(A2)', 'ISTEXT(A4)', 'ISERROR(A6)', 'ISNA(A6)',
    'SUM(A1:A5)', 'SUM("2",TRUE)', 'SUM(A6:A6)', 'MIN("2",TRUE)', 'MAX(A1:A5)', 'AVERAGE(A1:A5)', 'AVERAGE(A5:A5)',
    'COUNT("2",TRUE)', 'COUNT(A1:A5)', 'COUNTA(A1:A5)', 'COUNTBLANK(A1:A5)', 'SUM(,)', 'COUNT(,)', 'COUNTA(,)',
    'ABS(-3)', 'INT(-1.2)', 'TRUNC(-1.2)', 'MOD(-5,3)', 'MOD(5,-3)', 'POWER(2,3)', 'SQRT(4)', 'LEN("中文Ab")',
    'ROW()', 'COLUMN()', 'ROWS(A1:B3)', 'COLUMNS(A1:B3)', 'Constant', 'ComputedName', 'SUM((A1:A2,A2:A3))', 'SUM(A:A)',
]
STANDARD_DIFFERENCES = {
    '"A"&TRUE': ({"type": "text", "value": "ATRUE"}, "Microsoft calculation-operators-and-precedence-in-excel: logical-to-text conversion example"),
    'SQRT(-1)': ({"type": "error", "value": "#NUM!"}, "ECMA-376 Part 1 18.17.7.296: negative square-root argument"),
    'TRUE=1': ({"type": "boolean", "value": False}, "Declared typed comparison: logical and numeric values remain distinct without arithmetic coercion; Excel/WPS unverified"),
    'AND(TRUE,"TRUE")': ({"type": "boolean", "value": True}, "Declared direct logical-argument conversion accepts TRUE/FALSE text; reference text remains ignored; Excel/WPS unverified"),
    'SUM(A1:A5)': ({"type": "number", "value": 2}, "ECMA-376 Part 1 18.17.7.306: reference logical/text values ignored"),
    'SUM("2",TRUE)': ({"type": "number", "value": 3}, "ECMA-376 Part 1 18.17.7.306: direct logical/numeric text arguments counted"),
    'MIN("2",TRUE)': ({"type": "number", "value": 1}, "ECMA-376 Part 1 18.17.7.212: direct logical/numeric text arguments counted"),
    'AVERAGE(A1:A5)': ({"type": "number", "value": 2}, "ECMA-376 Part 1 18.17.7.18: reference logical/text values ignored"),
    'COUNT(A1:A5)': ({"type": "number", "value": 1}, "ECMA-376 Part 1 18.17.7.52: reference logical/text values ignored"),
    'SUM((A1:A2,A2:A3))': ({"type": "number", "value": 2}, "ECMA-376 Part 1 18.17.7.306: reference logical/text values ignored, union multiplicity retained"),
}


def generate(cached=True):
    book = native.fixture.Book(); sheet = book.sheet("Scalar")
    for address, value in [("A1", 2), ("A2", "2"), ("A3", True), ("A7", 0), ("A8", "x")]: sheet.add(address, value)
    sheet.add("A4", "" if cached else None, kind="str", formula='""', formula_attrs={"ca": "1"})
    sheet.add("A6", "#N/A", kind="e")
    for address, formula in [("C1", "A1+1"), ("C2", "C1*2")]: sheet.add(address, 999 if cached else None, formula=formula, formula_attrs={"ca": "1"})
    for index, formula in enumerate(FORMULAS, 1): sheet.add("E" + str(index), 999 if cached else None, formula=formula, formula_attrs={"ca": "1"})
    book.names = [({"name": "Constant"}, "7"), ({"name": "ComputedName"}, "SUM($A$1,$A$7)")]
    parts = book.finish(); parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return native.fixture.zipped(parts)


def observe(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read("xl/worksheets/sheet1.xml")); shared = []
        if "xl/sharedStrings.xml" in archive.namelist():
            shared = [''.join(item.itertext()) for item in ET.fromstring(archive.read("xl/sharedStrings.xml"))]
    cells = {cell.get("r"): cell for cell in root.findall("{*}sheetData/{*}row/{*}c")}
    results = []
    for index, formula in enumerate(FORMULAS, 1):
        address = "E" + str(index); cell = cells[address]; kind = cell.get("t", "n"); value = cell.findtext("{*}v")
        if kind == "b": result = {"type": "boolean", "value": value == "1"}
        elif kind == "e": result = {"type": "error", "value": value}
        elif kind in ["s", "str", "inlineStr"]:
            result = {"type": "text", "value": shared[int(value)] if kind == "s" else cell.findtext("{*}is/{*}t", "") if kind == "inlineStr" else value or ""}
        else: result = {"type": "number", "value": float(value)}
        expected, basis = STANDARD_DIFFERENCES.get(formula, (result, "Pinned Calc observation; Excel/WPS unverified"))
        results.append({"address": address, "expression": "=" + formula, "value": result, "expected": expected, "basis": basis,
                        "matchesProducer": expected == result})
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    cache = ROOT / ".cache/xlsx-scalar-semantics"; cache.mkdir(exist_ok=True)
    data = generate(); reference = generate(cached=False)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); path = work / "scalar.xlsx"; path.write_bytes(reference)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(output), str(path)])
        observations = observe(output / "scalar.xlsx")
    assert observations[0]["value"] == {"type": "number", "value": 7}, "Producer did not calculate the reference input"
    assert observations[8]["value"] == {"type": "error", "value": "#DIV/0!"}
    assert observations[21]["value"] == {"type": "number", "value": 3}
    report = {"scope": "Pinned Calc scalar operators, coercion, references and selected basic functions. Deliberately wrong source caches are not expected current results. Full calculation, Excel/WPS and later function families are separate gates.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "referenceInputSha256": sha256(reference).hexdigest(),
              "referenceInput": "Same formulas and constants with formula caches omitted; pinned Calc retains deliberately wrong caches despite native recalc flags, so those caches are never reference results.", "cases": observations}
    if args.record: FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == data, "Independent scalar source changed"
        assert json.loads(REPORT.read_text()) == report, "Native scalar results changed"
    print(f"Verified {len(observations)} independent scalar formula results with pinned Calc")


if __name__ == "__main__": main()
