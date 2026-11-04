"""Record fixed-font native cell alignment and text bounds without SDK output.

The explicit C.UTF-8 reference environment is not an SDK locale default. These
PDF observations calibrate reading layout; date/locale policy remains separate.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("alignment_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-alignment-semantics.json"
FONTS = [("Liberation Sans", 11), ("Liberation Serif", 13), ("Noto Sans CJK SC", 12)]


def inputs():
    cases = []
    for rtl in [False, True]:
        for order in [0, 1, 2]:
            for kind, value in [("inlineStr", "Value"), ("inlineStr", "אבג"), ("inlineStr", "123 אבג"), ("n", 123), ("n", -123), ("n", -123.5), ("b", True), ("e", "#DIV/0!")]:
                cases.append({"group": "general", "rightToLeft": rtl, "readingOrder": order, "horizontal": "general", "vertical": "bottom", "font": 0, "kind": kind, "value": value, "indent": 0})
        for horizontal in ["left", "center", "right"]:
            for value in ["Value", "אבג"]:
                cases.append({"group": "explicit", "rightToLeft": rtl, "readingOrder": 0, "horizontal": horizontal, "vertical": "bottom", "font": 0, "kind": "inlineStr", "value": value, "indent": 0})
        for font in range(3):
            for vertical in ["top", "center", "bottom"]:
                cases.append({"group": "vertical", "rightToLeft": rtl, "readingOrder": 0, "horizontal": "left", "vertical": vertical, "font": font, "kind": "inlineStr", "value": "汉字" if font == 2 else "Value", "indent": 0})
        for font in [0, 1]:
            for horizontal in ["left", "right"]:
                for indent in [1, 2]:
                    cases.append({"group": "indent", "rightToLeft": rtl, "readingOrder": 0, "horizontal": horizontal, "vertical": "bottom", "font": font, "kind": "inlineStr", "value": "Value", "indent": indent})
    return cases


def generate(cases):
    book = native.fixture.Book(); ns = "{" + native.fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"]); fonts = styles.find(ns + "fonts"); borders = styles.find(ns + "borders"); cells = styles.find(ns + "cellXfs")
    fonts.clear(); borders.clear(); cells.clear(); base = {"fontId": "0", "fillId": "0", "numFmtId": "0", "xfId": "0"}
    for name, size in FONTS:
        font = ET.SubElement(fonts, ns + "font"); ET.SubElement(font, ns + "sz", {"val": str(size)}); ET.SubElement(font, ns + "name", {"val": name}); ET.SubElement(font, ns + "color", {"rgb": "FF000000"})
    ET.SubElement(borders, ns + "border"); border = ET.SubElement(borders, ns + "border")
    for side in ["left", "right", "top", "bottom"]: ET.SubElement(ET.SubElement(border, ns + side, {"style": "thin"}), ns + "color", {"rgb": "FFFF0000"})
    ET.SubElement(cells, ns + "xf", {**base, "borderId": "0"})
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$C$3"; sheet.add("A1", "Case" + str(index)); sheet.row_props[2] = {"ht": 60, "customHeight": 1}
        if case["rightToLeft"]: sheet.before = sheet.before.replace('workbookViewId="0"', 'workbookViewId="0" rightToLeft="1"')
        node = ET.SubElement(cells, ns + "xf", {**base, "fontId": str(case["font"]), "borderId": "1", "applyBorder": "1", "applyAlignment": "1", "applyFont": "1"})
        ET.SubElement(node, ns + "alignment", {key: str(case[key]) for key in ["horizontal", "vertical", "readingOrder", "indent"]})
        sheet.add("B2", case["value"], kind=case["kind"], style=index)
    fonts.set("count", str(len(fonts))); borders.set("count", str(len(borders))); cells.set("count", str(len(cells)))
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode"); return native.fixture.zipped(book.finish())


def observe(pdf, cases, work):
    bbox = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"]))
    pages = bbox.findall(".//{http://www.w3.org/1999/xhtml}page"); assert len(pages) == len(cases)
    for index, (page, case) in enumerate(zip(pages, cases), 1):
        svg = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
        vertical, horizontal = set(), set()
        for path in ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"):
            if path.get("stroke") != "rgb(100%, 0%, 0%)": continue
            a, b, c, d, e, f = [float(value) for value in re.fullmatch(r"matrix\(([^)]+)\)", path.get("transform"))[1].split(",")]; assert b == c == 0
            points = re.fullmatch(r"M ([\d.-]+) ([\d.-]+) L ([\d.-]+) ([\d.-]+)\s*", path.get("d")); assert points
            x1, y1, x2, y2 = map(float, points.groups())
            if x1 == x2: vertical.add(round(x1 * a + e, 6))
            elif y1 == y2: horizontal.add(round(y1 * d + f, 6))
            else: raise AssertionError("Cell border is not axis aligned")
        assert len(vertical) == len(horizontal) == 2
        left, right = sorted(vertical); top, bottom = sorted(horizontal); words = []
        for word in page.findall("{http://www.w3.org/1999/xhtml}word"):
            x1, y1, x2, y2 = [float(word.get(key)) for key in ["xMin", "yMin", "xMax", "yMax"]]
            if left <= (x1 + x2) / 2 <= right and top <= (y1 + y2) / 2 <= bottom:
                words.append({"text": word.text, "bounds": [x1, y1, x2, y2]})
        assert words, case
        bounds = [min(word["bounds"][0] for word in words), min(word["bounds"][1] for word in words), max(word["bounds"][2] for word in words), max(word["bounds"][3] for word in words)]
        case["cellBoundsPt"] = [left, top, right, bottom]; case["pdfWords"] = words
        case["textGapsPt"] = [round(bounds[0] - left, 6), round(bounds[1] - top, 6), round(right - bounds[2], 6), round(bottom - bounds[3], 6)]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "alignment.xlsx"; cases = inputs(); data = generate(cases); source.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
        observe(source.with_suffix(".pdf"), cases, work)
    report = {"scope": "Pinned Calc point-space PDF text and cell-border bounds for General, explicit horizontal/vertical alignment, reading direction and indentation. This is not an SDK locale default, complete layout acceptance or Excel/WPS observation.",
              "archiveSha256": native.ARCHIVE_HASH, "environment": environment, "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"},
              "fonts": [{"family": name, "sizePt": size} for name, size in FONTS], "sourceSha256": sha256(data).hexdigest(), "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native alignment observations changed"
    print(f"Verified {len(cases)} native cell alignment and text-bound observations")


if __name__ == "__main__": main()
