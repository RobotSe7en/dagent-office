"""Observe which covered cells supply merged borders in the pinned Calc build.

These intentionally inconsistent perimeter styles expose producer behavior;
they do not establish Excel/WPS precedence or an SDK merge-rendering policy.
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
spec = importlib.util.spec_from_file_location("merged_border_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-merged-border-semantics.json"
SIDES = ["left", "right", "top", "bottom"]


def generate():
    book = native.fixture.Book(); namespace = "{" + native.fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"]); borders = styles.find(namespace + "borders"); cells = styles.find(namespace + "cellXfs")
    borders.clear(); cells.clear(); base = {"fontId": "0", "fillId": "0", "numFmtId": "0", "xfId": "0"}
    ET.SubElement(borders, namespace + "border"); ET.SubElement(cells, namespace + "xf", {**base, "borderId": "0"})
    def edge(sides, color="FF0000", kind="thin"):
        index = len(borders); border = ET.SubElement(borders, namespace + "border")
        for side in sides:
            ET.SubElement(ET.SubElement(border, namespace + side, {"style": kind}), namespace + "color", {"rgb": "FF" + color})
        ET.SubElement(cells, namespace + "xf", {**base, "borderId": str(index), "applyBorder": "1"}); return index
    kinds = {side: edge([side]) for side in SIDES}; cases = []
    for rtl in [False, True]:
        for row in range(1, 4):
            for column in "ABC":
                for side in SIDES:
                    address = column + str(row); name = ("R" if rtl else "L") + "-" + address + "-" + side
                    sheet = book.sheet(name); sheet.print_area = "$A$1:$E$4"
                    if rtl: sheet.before = sheet.before.replace('workbookViewId="0"', 'workbookViewId="0" rightToLeft="1"')
                    sheet.after = '<mergeCells count="1"><mergeCell ref="A1:C3"/></mergeCells>'; sheet.add("E1", name)
                    for source_row in range(1, 4):
                        sheet.row_props[source_row] = {"ht": 30, "customHeight": 1}
                        for source_column in "ABC":
                            cell = source_column + str(source_row)
                            sheet.add(cell, "anchor" if cell == "A1" else None, style=kinds[side] if cell == address else 0)
                    cases.append({"name": name, "rightToLeft": rtl, "sourceCell": address, "sourceSide": side})
        for side, outside, incoming in [("top", "C1", "bottom"), ("bottom", "C5", "top"), ("left", "A3", "right"), ("right", "E3", "left")]:
            name = ("R" if rtl else "L") + "-neighbor-" + side; sheet = book.sheet(name); sheet.print_area = "$A$1:$F$6"
            if rtl: sheet.before = sheet.before.replace('workbookViewId="0"', 'workbookViewId="0" rightToLeft="1"')
            sheet.after = '<mergeCells count="1"><mergeCell ref="B2:D4"/></mergeCells>'; sheet.add("F1", name)
            for source_row in range(2, 5):
                sheet.row_props[source_row] = {"ht": 30, "customHeight": 1}
                for source_column in "BCD":
                    cell = source_column + str(source_row); source_sides = []
                    if source_row == 2: source_sides.append("top")
                    if source_row == 4: source_sides.append("bottom")
                    if source_column == "B": source_sides.append("left")
                    if source_column == "D": source_sides.append("right")
                    sheet.add(cell, "anchor" if cell == "B2" else None, style=edge(source_sides))
            sheet.add(outside, style=edge([incoming], "00FF00", "medium"))
            cases.append({"name": name, "rightToLeft": rtl, "outsideCell": outside, "outsideSide": incoming, "mergedSide": side})
    borders.set("count", str(len(borders))); cells.set("count", str(len(cells))); book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    return native.fixture.zipped(book.finish()), cases


def observe(pdf, cases, work):
    pages = int(re.search(r"^Pages:\s+(\d+)$", native.shared.run(["pdfinfo", str(pdf)]), re.M)[1]); assert pages == len(cases)
    for page, case in enumerate(cases, 1):
        svg = work / f"page-{page}.svg"; native.shared.run(["pdftocairo", "-f", str(page), "-l", str(page), "-svg", str(pdf), str(svg)])
        strokes = []
        for path in ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"):
            color = path.get("stroke"); owner = {"rgb(100%, 0%, 0%)": "merged", "rgb(0%, 100%, 0%)": "neighbor"}.get(color)
            if not owner: continue
            a, b, c, d, e, f = [float(value) for value in re.fullmatch(r"matrix\(([^)]+)\)", path.get("transform"))[1].split(",")]
            assert b == c == 0 and abs(abs(a) - abs(d)) < .000001
            points = re.fullmatch(r"M ([\d.-]+) ([\d.-]+) L ([\d.-]+) ([\d.-]+)\s*", path.get("d")); assert points
            x1, y1, x2, y2 = map(float, points.groups()); assert x1 == x2 or y1 == y2
            strokes.append({"owner": owner, "from": [round(x1 * a + e, 6), round(y1 * d + f, 6)], "to": [round(x2 * a + e, 6), round(y2 * d + f, 6)],
                            "widthPt": round(float(path.get("stroke-width")) * abs(a), 6)})
        case["pdfStrokes"] = strokes
        if "sourceCell" in case:
            origin = {"top": "A1", "left": "A1", "right": "C1", "bottom": "A3"}[case["sourceSide"]]
            # The native RTL export can repeat an identical stroke. Preserve
            # those PDF records, but count unique geometry for edge presence.
            unique = {json.dumps(stroke, sort_keys=True) for stroke in strokes}
            assert len(unique) == (1 if case["sourceCell"] == origin else 0), case
        else:
            assert any(stroke["owner"] == "neighbor" for stroke in strokes), case
            assert any(stroke["owner"] == "merged" for stroke in strokes), case
    return pages


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "merged.xlsx"; data, cases = generate(); source.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
        pages = observe(source.with_suffix(".pdf"), cases, work)
    report = {"scope": "Pinned Calc observations for isolated borders on a 3x3 merged range and stronger outside neighbors, in LTR/RTL. Inconsistent source perimeter styles are intentional; this does not establish Excel/WPS precedence or SDK rendering policy.",
              "archiveSha256": native.ARCHIVE_HASH, "environment": environment, "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"},
              "sourceSha256": sha256(data).hexdigest(), "pdfPages": pages, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native merged-border observations changed"
    print(f"Verified {len(cases)} native merged-border PDF observations")


if __name__ == "__main__": main()
