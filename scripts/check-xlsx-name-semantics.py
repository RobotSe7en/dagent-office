"""Probe defined-name origins and scope using independent ZIP/XML and pinned Calc.

ODS is an external semantic observation, never an SDK conversion dependency.
"""
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
BUILD = "LibreOffice 25.8.1.1 54047653041915e595ad4e45cccea684809c77b5"
ARCHIVE_HASH = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
REPORT = ROOT / "docs/evidence/xlsx-name-semantics.json"
T = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
O = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    result = importlib.util.module_from_spec(spec); spec.loader.exec_module(result); return result


fixture = module("independent_name_inputs", "scripts/generate-xlsx-fixtures.py")
shared = module("office_name_reference", "scripts/pptx-reference.py")


def generate(active, active_sheet):
    book = fixture.Book()
    for index, name in enumerate(["Data", "Second", "Third"]):
        sheet = book.sheet(name)
        sheet.before = sheet.before.replace('<sheetView workbookViewId="0"/>', f'<sheetView workbookViewId="0"><selection activeCell="{active}" sqref="{active}"/></sheetView>')
        for cell, value in [("A1", 10), ("A2", [20, 222, 2222][index]), ("A3", 30), ("A4", [44, 444, 4444][index]), ("D1", 400)]: sheet.add(cell, value)
        for cell, formula in [("C2", "Relative"), ("C3", "Relative"), ("D3", "RelativeColumn"), ("C4", "Unqualified"), ("D4", "Scoped")]: sheet.add(cell, formula=formula, formula_attrs={"ca": "1"})
    book.names = [({"name": "Relative"}, "Data!$A1"), ({"name": "RelativeColumn"}, "Data!A$1"), ({"name": "Unqualified"}, "$A1"),
                  ({"name": "Scoped"}, "100"), ({"name": "Scoped", "localSheetId": 0}, "10"), ({"name": "Scoped", "localSheetId": 1}, "20")]
    parts = book.finish(); parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('activeTab="0"', f'activeTab="{active_sheet}"').replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return fixture.zipped(parts)


def observe(path):
    with ZipFile(path) as archive: root = ET.fromstring(archive.read("content.xml"))
    names, sheets = [], {}
    parents = {child: parent for parent in root.iter() for child in parent}
    for node in root.iter():
        if node.tag not in [f"{{{T}}}named-range", f"{{{T}}}named-expression"]: continue
        name = node.get(f"{{{T}}}name")
        if name.startswith("_xlnm."): continue
        parent = node; scope = None
        while parent in parents:
            parent = parents[parent]
            if parent.tag == f"{{{T}}}table": scope = parent.get(f"{{{T}}}name"); break
        names.append({"name": name, "scope": scope, "origin": node.get(f"{{{T}}}base-cell-address"),
                      "reference": node.get(f"{{{T}}}cell-range-address"), "expression": node.get(f"{{{T}}}expression")})
    for table in root.iter(f"{{{T}}}table"):
        cells = {}; row_index = 0
        for row in table.findall(f"{{{T}}}table-row"):
            column = 0
            for cell in row:
                if cell.tag not in [f"{{{T}}}table-cell", f"{{{T}}}covered-table-cell"]: continue
                formula = cell.get(f"{{{T}}}formula")
                if formula is not None:
                    assert cell.get(f"{{{O}}}value-type") == "float", "A named formula produced an error or nonnumeric result"
                    cells[fixture.column(column + 1) + str(row_index + 1)] = {"formula": formula, "value": cell.get(f"{{{O}}}value")}
                column += int(cell.get(f"{{{T}}}number-columns-repeated", "1"))
            row_index += int(row.get(f"{{{T}}}number-rows-repeated", "1"))
        sheets[table.get(f"{{{T}}}name")] = cells
    return {"names": sorted(names, key=lambda n: (n["scope"] or "", n["name"])), "sheets": sheets}


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert shared.run(["libreoffice", "--version"]) == BUILD
    cases = []; cache = ROOT / ".cache/xlsx-name-semantics"; cache.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for active, sheet_index in [("A1", 0), ("H8", 1)]:
            source = work / (active + ".xlsx"); source.write_bytes(generate(active, sheet_index))
            shared.run(["libreoffice", profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
            observed = observe(source.with_suffix(".ods")); assert len(observed["names"]) == 6
            assert all(name["origin"].endswith(".$A$1") for name in observed["names"])
            for index, name in enumerate(["Data", "Second", "Third"]):
                expected = {"C2": ("Relative", "20"), "C3": ("Relative", "30"), "D3": ("RelativeColumn", "400"), "C4": ("Unqualified", str([44, 444, 4444][index])), "D4": ("Scoped", str([10, 20, 100][index]))}
                assert observed["sheets"][name] == {cell: {"formula": "of:=" + formula, "value": value} for cell, (formula, value) in expected.items()}, observed["sheets"][name]
            cases.append({"activeCell": active, "activeSheetIndex": sheet_index, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "observed": observed})
    report = {"scope": "Relative defined-name origin, caller-sheet binding and local shadowing only. Excel/WPS and SDK calculation remain unverified.", "office": BUILD, "archiveSha256": ARCHIVE_HASH, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native defined-name semantics changed"
    print("Verified 2 independent name-origin probes, 3 sheets and 30 named-formula results with pinned Calc")


if __name__ == "__main__": main()
