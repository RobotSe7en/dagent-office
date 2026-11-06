"""Observe regular SpreadsheetML borders through independent ODS and PDF output.

This records the pinned Calc renderer, including its loss of slantDashDot.
It is not an Excel/WPS observation or a complete SDK border-rendering gate.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("native_border_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-border-semantics.json"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
KINDS = ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"]
PAIRS = [(a, b) for a in KINDS for b in KINDS]
BASE = {"fontId": "0", "fillId": "0", "numFmtId": "0", "xfId": "0"}


def generate(axis=None, rtl=False):
    book = native.fixture.Book(); sheet = book.sheet("Borders"); namespace = "{" + native.fixture.S + "}"
    if rtl: sheet.before = sheet.before.replace('workbookViewId="0"', 'workbookViewId="0" rightToLeft="1"')
    sheet.print_area = f"$A$1:$C${len(PAIRS) * 3}" if axis else "$A$1:$A$14"
    sheet.page = '<pageMargins left="0.5" right="0.5" top="0.5" bottom="0.5" header="0.2" footer="0.2"/><pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/>'
    styles = ET.fromstring(book.parts["xl/styles.xml"]); borders = styles.find(namespace + "borders"); cells = styles.find(namespace + "cellXfs")
    borders.clear(); cells.clear(); ET.SubElement(borders, namespace + "border"); ET.SubElement(cells, namespace + "xf", {**BASE, "borderId": "0"})
    def edge(kind, side, color):
        index = len(borders)
        border = ET.SubElement(borders, namespace + "border"); value = ET.SubElement(border, namespace + side, {"style": kind})
        ET.SubElement(value, namespace + "color", {"rgb": "FF" + color})
        ET.SubElement(cells, namespace + "xf", {**BASE, "borderId": str(index), "applyBorder": "1"})
        return index
    if axis:
        for index, pair in enumerate(PAIRS, 1):
            row = (index - 1) * 3 + 1; sheet.add("A" + str(row), f"{index}: {pair[0]}/{pair[1]}")
            for owner, kind in enumerate(pair):
                side = ("bottom", "top")[owner] if axis == "row" else ("right", "left")[owner]
                color = f"{index:02X}0000" if owner == 0 else f"00{index:02X}00"
                address = "B" + str(row + owner) if axis == "row" else ("B", "C")[owner] + str(row)
                sheet.add(address, style=edge(kind, side, color))
    else:
        for row, kind in enumerate(KINDS, 1):
            sheet.add("A" + str(row), kind, style=edge(kind, "bottom", f"{row:02X}0000")); sheet.row_props[row] = {"ht": 30, "customHeight": 1}
    borders.set("count", str(len(borders))); cells.set("count", str(len(cells)))
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    return native.fixture.zipped(book.finish())


def observe_styles(path):
    with ZipFile(path) as archive: roots = [ET.fromstring(archive.read(part)) for part in ["styles.xml", "content.xml"]]
    definitions = {node.get(f"{{{STYLE}}}name"): node for root in roots for node in root.iter(f"{{{STYLE}}}style")}
    column_default = roots[1].find(f".//{{{native.T}}}table-column").get(f"{{{native.T}}}default-cell-style-name")
    result = []
    for cell in roots[1].iter(f"{{{native.T}}}table-cell"):
        if not cell.get(f"{{{native.O}}}value-type"): continue
        kind = "".join(cell.itertext()); node = definitions[cell.get(f"{{{native.T}}}style-name") or column_default].find(f"{{{STYLE}}}table-cell-properties")
        result.append({"kind": kind, "nativeOdsBorder": node.get(f"{{{FO}}}border-bottom") or node.get(f"{{{FO}}}border"),
                       "nativeDoubleWidths": node.get(f"{{{STYLE}}}border-line-width-bottom")})
    assert [sample["kind"] for sample in result] == KINDS
    return result


def observe_edges(pdf, work):
    pages = int(re.search(r"^Pages:\s+(\d+)$", native.shared.run(["pdfinfo", str(pdf)]), re.M)[1]); assert 1 <= pages <= 20
    owners = [set() for _ in PAIRS]
    for page in range(1, pages + 1):
        svg = work / f"{pdf.stem}-{page}.svg"
        native.shared.run(["pdftocairo", "-f", str(page), "-l", str(page), "-svg", str(pdf), str(svg)])
        for path in ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"):
            stroke = path.get("stroke", ""); match = re.fullmatch(r"rgb\(([\d.]+)%, ([\d.]+)%, ([\d.]+)%\)", stroke)
            if not match: continue
            red, green, blue = [int(float(value) * 2.55 + .5) for value in match.groups()]
            if red and not green and not blue: owners[red - 1].add("first")
            elif green and not red and not blue: owners[green - 1].add("second")
            elif red or green or blue: raise AssertionError("Unexpected PDF border color " + stroke)
    bbox = native.shared.run(["pdftotext", "-f", "1", "-l", "1", "-bbox", str(pdf), "-"])
    leading_x = float(re.search(r'<word xMin="([\d.]+)"[^>]*>1:', bbox)[1])
    return pages, leading_x, [sorted(value) for value in owners]


def observe_strokes(pdf, work):
    """Read actual device-independent PDF strokes, including page transforms."""
    svg = work / "kinds.svg"; native.shared.run(["pdftocairo", "-f", "1", "-l", "1", "-svg", str(pdf), str(svg)])
    result = {kind: [] for kind in KINDS}
    for path in ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"):
        color = re.fullmatch(r"rgb\(([\d.]+)%, 0%, 0%\)", path.get("stroke", ""))
        if not color: continue
        row = int(float(color[1]) * 2.55 + .5); assert 1 <= row <= len(KINDS)
        transform = [float(value) for value in re.fullmatch(r"matrix\(([^)]+)\)", path.get("transform"))[1].split(",")]
        a, b, c, d, e, f = transform; assert b == c == 0 and abs(abs(a) - abs(d)) < .000001
        points = re.fullmatch(r"M ([\d.-]+) ([\d.-]+) L ([\d.-]+) ([\d.-]+)\s*", path.get("d")); assert points
        x1, y1, x2, y2 = map(float, points.groups()); assert y1 == y2 and x2 > x1
        result[KINDS[row - 1]].append({"from": [round(x1 * a + e, 6), round(y1 * d + f, 6)], "to": [round(x2 * a + e, 6), round(y2 * d + f, 6)],
            "widthPt": round(float(path.get("stroke-width")) * abs(a), 6), "dashPt": [round(float(value) * abs(a), 6) for value in path.get("stroke-dasharray", "").split()],
            "cap": path.get("stroke-linecap")})
    assert len(result["none"]) == 0 and len(result["double"]) == 2
    assert all(len(result[kind]) == 1 for kind in KINDS if kind not in ["none", "double", "slantDashDot"])
    # Calc can duplicate its substituted slantDashDot stroke; retain that fact.
    assert result["slantDashDot"] and all(value == result["slantDashDot"][0] for value in result["slantDashDot"])
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); profile = "-env:UserInstallation=" + (work / "profile").as_uri(); source = work / "kinds.xlsx"; source.write_bytes(generate())
        native.shared.run(["libreoffice", profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        native.shared.run(["libreoffice", profile, "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
        kinds = {"sourceSha256": sha256(source.read_bytes()).hexdigest(), "samples": observe_styles(source.with_suffix(".ods")), "pdfStrokes": observe_strokes(source.with_suffix(".pdf"), work)}; cases = []
        for axis, rtl in [("row", False), ("column", False), ("column", True)]:
            source = work / f"{axis}-{'rtl' if rtl else 'ltr'}.xlsx"; source.write_bytes(generate(axis, rtl))
            native.shared.run(["libreoffice", profile, "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
            pages, leading_x, owners = observe_edges(source.with_suffix(".pdf"), work)
            # Assert the inferred rule against actual PDF colors, not XLSX style
            # records or an SDK renderer. Equal widths/patterns choose physical
            # bottom/right; double ranks above medium and below thick.
            rank = {"none": 0, "hair": 1, "thin": 2, "dashed": 2, "dotted": 2, "dashDot": 2, "dashDotDot": 2, "medium": 3, "mediumDashed": 3, "mediumDashDot": 3, "mediumDashDotDot": 3, "slantDashDot": 3, "double": 4, "thick": 5}
            for (a, b), observed in zip(PAIRS, owners):
                expected = [] if a == b == "none" else ["first" if rank[a] > rank[b] or rank[a] == rank[b] and rtl else "second"]
                assert observed == expected, (axis, rtl, a, b, observed)
            assert leading_x > 400 if rtl else leading_x < 100
            cases.append({"axis": axis, "rightToLeft": rtl, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "pdfPages": pages, "leadingLabelXPt": leading_x,
                          "samples": [{"first": a, "second": b, "owners": observed} for (a, b), observed in zip(PAIRS, owners)]})
    report = {"scope": "Pinned Calc regular-border widths and 588 shared-edge PDF color observations. Conditional formats, merged perimeters and Excel/WPS remain unverified. Calc maps slantDashDot to fine-dashed; this is a producer difference, not proof of faithful slanted-border rendering.",
              "archiveSha256": native.ARCHIVE_HASH, "environment": environment, "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "kinds": kinds, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native border observations changed"
    print(f"Verified {len(KINDS)} native border records and {len(PAIRS) * len(cases)} PDF shared-edge selections")


if __name__ == "__main__": main()
