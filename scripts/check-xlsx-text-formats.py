"""Record plain/rich text-format behavior from independent XLSX and pinned Calc.

Rich text and format precedence differ in this producer; the observations do
not claim Excel/WPS behavior and do not decide the SDK's product policy.
"""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("text_formats", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-text-formats.xlsx"
REPORT = ROOT / "docs/evidence/xlsx-text-formats.json"
CASES = [
    {"value": "001😀", "code": '"ID "@'},
    {"value": "A/B", "code": '@" / "@'},
    {"value": "AB", "code": '[Red]"ID "@'},
    {"value": "AB", "code": '0;0;0;"fixed"'},
    {"value": "AB", "code": ';;;'},
    {"value": "AB", "code": '@_)*.'},
    {"value": "AB", "code": '*x@'},
    {"value": "AB", "code": '@*x'},
    {"value": "AB", "code": '*x@*.'},
    {"value": "A😀", "code": '"ID "@', "rotation": 255},
    {"value": "AB", "code": '[Red]"ID "@', "rich": True},
    {"value": "AB", "code": '"ID "@', "rich": True},
    {"value": "AB", "code": '@"/"@', "rich": True},
]


def generate():
    native.CASES = [(case["value"], case["code"]) for case in CASES]
    with ZipFile(BytesIO(native.generate(False))) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    ns = "{" + native.fixture.S + "}"
    sheet = ET.fromstring(parts["xl/worksheets/sheet1.xml"])
    columns = ET.Element(ns + "cols"); ET.SubElement(columns, ns + "col", {"min": "1", "max": "1", "width": "40", "customWidth": "1"})
    sheet.insert(list(sheet).index(sheet.find(ns + "sheetData")), columns)
    styles = ET.fromstring(parts["xl/styles.xml"]); formats = styles.find(ns + "cellXfs")
    for index, cell in enumerate(sheet.findall(".//" + ns + "c")):
        case = CASES[index]
        if case.get("rotation"):
            ET.SubElement(formats[index + 1], ns + "alignment", {"textRotation": str(case["rotation"])})
            row = sheet.find(ns + "sheetData")[index]; row.set("ht", "100"); row.set("customHeight", "1")
        if case.get("rich"):
            for child in list(cell): cell.remove(child)
            cell.set("t", "inlineStr"); value = ET.SubElement(cell, ns + "is")
            for text, color in [("A", "FF0000FF"), ("B", "FF00AA00")]:
                run = ET.SubElement(value, ns + "r"); properties = ET.SubElement(run, ns + "rPr")
                ET.SubElement(properties, ns + "color", {"rgb": color}); ET.SubElement(run, ns + "t").text = text
    parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    parts["xl/worksheets/sheet1.xml"] = ET.tostring(sheet, encoding="unicode")
    return native.fixture.zipped(parts)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    data = generate()
    if args.record: FIXTURE.write_bytes(data)
    else: assert FIXTURE.read_bytes() == data
    with TemporaryDirectory(dir=ROOT / ".cache") as directory:
        work = Path(directory); path = work / "text-formats.xlsx"; path.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(path)])
        observed = native.observe(path.with_suffix(".ods")); assert len(observed) == len(CASES)
    report = {"scope": __doc__.strip(), "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"},
              "samples": [{**case, "text": value} for case, value in zip(CASES, observed)]}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native text-format observations changed"
    print(f"Verified {len(CASES)} native plain/rich text-format observations")


if __name__ == "__main__": main()
