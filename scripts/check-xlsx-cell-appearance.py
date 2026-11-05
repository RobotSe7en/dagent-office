"""Independent advanced cell appearance sources and pinned Calc observations.

Native output is evidence of producer behavior, not a replacement for the
original pattern, gradient, border or text-rotation source.
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
spec = importlib.util.spec_from_file_location("appearance_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-cell-appearance.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-cell-appearance.xlsx"
N = "{" + native.fixture.S + "}"
PATTERNS = ["none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis", "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625"]


def inputs():
    result = [{"kind": "pattern", "pattern": pattern} for pattern in PATTERNS]
    result += [{"kind": "gradient", "type": "linear", "degree": degree} for degree in [0, 45, 90, 135, 270]]
    result += [{"kind": "gradient", "type": "path", **rectangle} for rectangle in [
        {"left": 0, "right": 0, "top": 0, "bottom": 0},
        {"left": .5, "right": .5, "top": .5, "bottom": .5},
        {"left": .2, "right": .8, "top": .3, "bottom": .7}]]
    result += [{"kind": "diagonal", "style": style, "up": up, "down": down} for style in ["thin", "double", "slantDashDot"] for up, down in [(True, False), (False, True), (True, True)]]
    result += [{"kind": "rotation", "rotation": rotation} for rotation in [0, 45, 90, 135, 180, 255]]
    return result


def generate(cases):
    book = native.fixture.Book(); styles = ET.fromstring(book.parts["xl/styles.xml"])
    fonts, fills, borders, formats = [styles.find(N + name) for name in ["fonts", "fills", "borders", "cellXfs"]]
    fonts.clear(); fills.clear(); borders.clear(); formats.clear()
    font = ET.SubElement(fonts, N + "font"); ET.SubElement(font, N + "sz", {"val": "12"}); ET.SubElement(font, N + "name", {"val": "Noto Sans CJK SC"}); ET.SubElement(font, N + "color", {"rgb": "FF000000"})
    for kind in ["none", "gray125"]: ET.SubElement(ET.SubElement(fills, N + "fill"), N + "patternFill", {"patternType": kind})
    ET.SubElement(borders, N + "border")
    base = {"fontId": "0", "fillId": "0", "borderId": "0", "numFmtId": "0", "xfId": "0"}
    ET.SubElement(formats, N + "xf", base)
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$C$3"; sheet.add("A1", "Case" + str(index)); sheet.row_props[2] = {"ht": 60, "customHeight": 1}
        props = {**base, "applyFont": "1", "applyFill": "1", "applyBorder": "1", "applyAlignment": "1"}
        if case["kind"] in ["pattern", "gradient"]:
            props["fillId"] = str(len(fills)); fill = ET.SubElement(fills, N + "fill")
            if case["kind"] == "pattern":
                node = ET.SubElement(fill, N + "patternFill", {"patternType": case["pattern"]})
                ET.SubElement(node, N + "fgColor", {"rgb": "FF204080"}); ET.SubElement(node, N + "bgColor", {"rgb": "FFFFE0A0"})
            else:
                node = ET.SubElement(fill, N + "gradientFill", {key: str(value) for key, value in case.items() if key != "kind"})
                for position, color in [(0, "FFFF0000"), (.5, "FF00FF00"), (1, "FF0000FF")]:
                    ET.SubElement(ET.SubElement(node, N + "stop", {"position": str(position)}), N + "color", {"rgb": color})
        props["borderId"] = str(len(borders)); border = ET.SubElement(borders, N + "border", {"diagonalUp": str(int(case.get("up", False))), "diagonalDown": str(int(case.get("down", False)))})
        for side in ["left", "right", "top", "bottom"]:
            ET.SubElement(ET.SubElement(border, N + side, {"style": "thin"}), N + "color", {"rgb": "FFCC00CC"})
        if case["kind"] == "diagonal": ET.SubElement(ET.SubElement(border, N + "diagonal", {"style": case["style"]}), N + "color", {"rgb": "FF0044AA"})
        format = ET.SubElement(formats, N + "xf", props)
        ET.SubElement(format, N + "alignment", {"horizontal": "center", "vertical": "center", "textRotation": str(case.get("rotation", 0))})
        sheet.add("B2", "预算 Ab" if case["kind"] == "rotation" else "", style=index)
    for table in [fonts, fills, borders, formats]: table.set("count", str(len(table)))
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode"); return native.fixture.zipped(book.finish())


def tree(node):
    if node is None: return None
    return {"tag": node.tag.split("}")[-1], **dict(sorted(node.attrib.items())), **({"children": [tree(child) for child in node]} if len(node) else {})}


def observe(path, count):
    with ZipFile(path) as archive:
        styles = ET.fromstring(archive.read("xl/styles.xml")); formats = styles.find(N + "cellXfs"); fills = styles.find(N + "fills"); borders = styles.find(N + "borders")
        result = []
        for index in range(1, count + 1):
            sheet = ET.fromstring(archive.read(f"xl/worksheets/sheet{index}.xml")); cell = sheet.find(f'.//{N}c[@r="B2"]'); assert cell is not None
            format = formats[int(cell.get("s", "0"))]
            result.append({"fill": tree(fills[int(format.get("fillId", "0"))]), "border": tree(borders[int(format.get("borderId", "0"))]), "alignment": tree(format.find(N + "alignment"))})
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    cases = inputs(); data = generate(cases); tables = {"fill": [], "border": [], "alignment": []}; indexes = {key: {} for key in tables}
    def intern(appearance):
        result = {}
        for key, definition in appearance.items():
            identity = json.dumps(definition, sort_keys=True)
            if identity not in indexes[key]: indexes[key][identity] = len(tables[key]); tables[key].append(definition)
            result[key] = indexes[key][identity]
        return result
    if args.record: FIXTURE.write_bytes(data)
    else: assert FIXTURE.read_bytes() == data, "Independent cell appearance input changed"
    reference = ROOT / ".cache/xlsx-cell-appearance-reference"; reference.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "appearance.xlsx"; source.write_bytes(data); producer = work / "producer"; producer.mkdir()
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(producer), str(source)])
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
        native.shared.run(["pdftocairo", "-png", "-r", "96", str(source.with_suffix(".pdf")), str(reference / "page")])
        pages = sorted(reference.glob("page-*.png")); assert len(pages) == len(cases)
        source_parts, producer_parts = observe(source, len(cases)), observe(producer / source.name, len(cases))
        for index, case in enumerate(cases):
            case["source"] = intern(source_parts[index]); case["producer"] = intern(producer_parts[index])
            case["pngSha256"] = sha256(pages[index].read_bytes()).hexdigest()
    report = {"scope": "Native cell appearance retention and fixed-font Calc page references. Producer changes are recorded separately; no SDK visual or Excel/WPS acceptance claim.",
              "sourceSha256": sha256(data).hexdigest(), "archiveSha256": native.ARCHIVE_HASH, "environment": environment,
              "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "appearanceTables": tables, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native cell appearance observations changed"
    print(f"Verified {len(cases)} advanced native cell appearance sources and references")


if __name__ == "__main__": main()
