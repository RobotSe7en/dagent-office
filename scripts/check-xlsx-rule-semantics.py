"""Probe native rule semantics using independent ZIP/XML and pinned Calc.

No SDK importer/exporter or spreadsheet library generates these inputs. ODS is
only an external acceptance observation, never an SDK conversion dependency.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZIP_DEFLATED, ZipInfo
import argparse
import json
import os
import subprocess
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
T = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
C = "urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
ARCHIVE_HASH = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
BUILD = "LibreOffice 25.8.1.1 54047653041915e595ad4e45cccea684809c77b5"
REPORT = ROOT / "docs/evidence/xlsx-rule-semantics.json"


def run(args):
    env = dict(os.environ)
    env["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    return subprocess.run([str(ROOT / ".cache/acceptance-office/bin/libreoffice"), *args], env=env, text=True, capture_output=True, check=True, timeout=60).stdout.strip()


def generate(destination, kind):
    with ZipFile(ROOT / "examples/xlsx/06-rules-validation.xlsx") as source, ZipFile(destination, "w", ZIP_DEFLATED) as output:
        for path in source.namelist():
            data = source.read(path)
            if path == "xl/worksheets/sheet1.xml":
                root = ET.fromstring(data)
                if kind == "dropdown":
                    validations = root.find(f"{{{S}}}dataValidations")
                    template = ET.tostring(validations[0]); validations.clear(); validations.set("count", "2")
                    for value, cell in [("0", "F2"), ("1", "G2")]:
                        node = ET.fromstring(template); node.set("sqref", cell); node.set("showDropDown", value); validations.append(node)
                else:
                    cf = root.findall(f"{{{S}}}conditionalFormatting")[1]
                    cf.set("sqref", "D5:D6 B2:B3" if kind == "reordered" else "D2 B5")
                    dv = root.find(f"{{{S}}}dataValidations")[-1]
                    dv.set("sqref", "H9:H10 E8" if kind == "reordered" else "H8 E9")
                data = ET.tostring(root, encoding="utf-8", xml_declaration=True)
            entry = ZipInfo(path, (1980, 1, 1, 0, 0, 0)); entry.compress_type = ZIP_DEFLATED; output.writestr(entry, data)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    assert run(["--version"]) == BUILD, "Unexpected external Office build"
    cache = ROOT / ".cache/xlsx-rule-semantics"; cache.mkdir(parents=True, exist_ok=True)
    cases = []
    with TemporaryDirectory(dir=cache) as temp:
        work = Path(temp); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["reordered", "cross-axis", "dropdown"]:
            source = work / (kind + ".xlsx"); generate(source, kind)
            run([profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
            with ZipFile(work / (kind + ".ods")) as output: root = ET.fromstring(output.read("content.xml"))
            validations = root.findall(f".//{{{T}}}content-validation")
            if kind == "dropdown":
                observed = {node.get(f"{{{T}}}base-cell-address"): node.get(f"{{{T}}}display-list") for node in validations}
                assert observed == {"Rules.F2": "unsorted", "Rules.G2": "no"}
            else:
                condition = root.findall(f".//{{{C}}}condition")[1]
                dv = validations[-1]
                observed = {"conditionalOrigin": condition.get(f"{{{C}}}base-cell-address"), "conditionalFormula": condition.get(f"{{{C}}}value"),
                            "validationOrigin": dv.get(f"{{{T}}}base-cell-address"), "validationFormula": dv.get(f"{{{T}}}condition")}
                cell = "H9" if kind == "reordered" else "H8"
                cf_origin = "Rules.B2" if kind == "reordered" else "Rules.B5"
                dv_reference = "H9" if kind == "reordered" else "H7"
                assert observed == {"conditionalOrigin": cf_origin, "conditionalFormula": "formula-is(AND([.B2]>=50;[.B2]<=150))",
                                    "validationOrigin": "Rules." + cell, "validationFormula": f"of:is-true-formula([.{dv_reference}]>0)"}, (kind, observed)
            cases.append({"case": kind, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "observed": observed})
    report = {"scope": "Native rule origin and dropdown semantics only; no Excel/WPS or full SDK reading acceptance claim.",
              "office": BUILD, "archiveSha256": ARCHIVE_HASH, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native rule semantics changed"
    print(f"Verified {len(cases)} independent native rule probes with pinned Calc")


if __name__ == "__main__": main()
