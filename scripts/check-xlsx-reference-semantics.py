"""Check independent union, intersection and qualified-name probes in pinned Calc.

The observed global nested-name binding is not generalized to Excel/WPS. ODS is
only external acceptance evidence, never an SDK conversion implementation.
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
REPORT = ROOT / "docs/evidence/xlsx-reference-semantics.json"
T = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
O = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
FORMULAS = ["Base", "Data!Base", "Data!LocalBase", "Chain", "Data!LocalChain", "SUM((A1:A2,A2:A3))", "SUM(A1:A3 A2:A4)", "SUM((A1:A2,A4:A5):A6)"]
FACTORS = {"Data": 1, "Second": 10, "Third": 100}


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    result = importlib.util.module_from_spec(spec); spec.loader.exec_module(result); return result


fixture = module("independent_reference_inputs", "scripts/generate-xlsx-fixtures.py")
shared = module("office_reference_probe", "scripts/pptx-reference.py")


def generate(order, active):
    book = fixture.Book()
    for name in order:
        sheet = book.sheet(name)
        for row in range(1, 7): sheet.add("A" + str(row), FACTORS[name] * row)
        for row, formula in enumerate(FORMULAS, 1): sheet.add("E" + str(row), formula=formula, formula_attrs={"ca": "1"})
    book.names = [({"name": "Base"}, "$A$2"), ({"name": "LocalBase", "localSheetId": order.index("Data")}, "$A$2"),
                  ({"name": "Chain"}, "Scoped"), ({"name": "LocalChain", "localSheetId": order.index("Data")}, "Scoped"),
                  ({"name": "Scoped"}, "$A$1"), ({"name": "Scoped", "localSheetId": order.index("Data")}, "$A$2"),
                  ({"name": "Scoped", "localSheetId": order.index("Second")}, "$A$3")]
    parts = book.finish()
    parts["xl/workbook.xml"] = parts["xl/workbook.xml"].replace('activeTab="0"', f'activeTab="{active}"').replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return fixture.zipped(parts)


def observe(path):
    with ZipFile(path) as archive: root = ET.fromstring(archive.read("content.xml"))
    result = {}
    for table in root.iter(f"{{{T}}}table"):
        cells = {}; row_index = 0
        for row in table.findall(f"{{{T}}}table-row"):
            column = 0
            for cell in row:
                if cell.tag not in [f"{{{T}}}table-cell", f"{{{T}}}covered-table-cell"]: continue
                formula = cell.get(f"{{{T}}}formula")
                if formula is not None:
                    numeric = cell.get(f"{{{O}}}value")
                    cells[fixture.column(column + 1) + str(row_index + 1)] = {"formula": formula, "value": numeric if numeric is not None else "".join(cell.itertext())}
                column += int(cell.get(f"{{{T}}}number-columns-repeated", "1"))
            row_index += int(row.get(f"{{{T}}}number-rows-repeated", "1"))
        result[table.get(f"{{{T}}}name")] = cells
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert shared.run(["libreoffice", "--version"]) == BUILD
    cache = ROOT / ".cache/xlsx-reference-semantics"; cache.mkdir(parents=True, exist_ok=True); cases = []
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for index, (order, active) in enumerate([(["Data", "Second", "Third"], 0), (["Data", "Second", "Third"], 1), (["Second", "Data", "Third"], 0)]):
            source = work / (str(index) + ".xlsx"); source.write_bytes(generate(order, active))
            shared.run(["libreoffice", profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
            observed = observe(source.with_suffix(".ods")); assert set(observed) == set(FACTORS)
            for name, factor in FACTORS.items():
                expected = [str(2 * factor), "#NAME?", "2", "2" if order[0] == "Data" else "30", "2", str(8 * factor), str(5 * factor), str(21 * factor)]
                assert {cell: entry["value"] for cell, entry in observed[name].items()} == {"E" + str(row): value for row, value in enumerate(expected, 1)}, observed[name]
            cases.append({"sheetOrder": order, "activeSheetIndex": active, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "observed": observed})
    report = {"scope": "Static reference operators and qualified local names. Calc binds a global nested unqualified name to a first-sheet local shadow; the SDK reports this uncertain global scope as unresolved. Excel/WPS and full SDK calculation remain unverified.", "office": BUILD, "archiveSha256": ARCHIVE_HASH, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native reference semantics changed"
    print("Verified 3 independent reference probes, 3 sheets and 72 formula results with pinned Calc")


if __name__ == "__main__": main()
