"""Record independent native rotated-cell outlines, fill and text placement.

Pinned Calc observations complement Microsoft's rotated-header instructions;
they do not constitute Excel/WPS verification. No SDK renderer is used.
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
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("rotated_cell_source", ROOT / "scripts/check-xlsx-cell-appearance.py")
source = importlib.util.module_from_spec(spec); spec.loader.exec_module(source)
native = source.native; N = source.N
REPORT = ROOT / "docs/evidence/xlsx-rotated-cells.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-rotated-cells.xlsx"


def inputs():
    result = [{"rotation": 0, "mode": "frame"}]
    result += [{"rotation": rotation, "mode": mode} for rotation in [30, 45, 60, 135] for mode in ["none", "frame", "top", "left", "diagonal"]]
    result += [{"rotation": rotation, "mode": "frame", "horizontal": horizontal, "vertical": "bottom"} for rotation in [45, 135] for horizontal in ["left", "right"]]
    result += [{"rotation": rotation, "mode": "frame", "vertical": vertical} for rotation in [45, 135] for vertical in ["top", "bottom"]]
    result += [{"rotation": rotation, "mode": "frame"} for rotation in [1, 89, 90, 91, 179, 180, 255]]
    return [{"kind": "rotation", "text": "Ab", "horizontal": "center", "vertical": "center", **item} for item in result]


def generate(cases):
    with ZipFile(BytesIO(source.generate(cases, 60))) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    styles = ET.fromstring(parts["xl/styles.xml"]); fills = styles.find(N + "fills"); formats = styles.find(N + "cellXfs"); borders = styles.find(N + "borders")
    fill = ET.SubElement(fills, N + "fill"); pattern = ET.SubElement(fill, N + "patternFill", {"patternType": "solid"})
    ET.SubElement(pattern, N + "fgColor", {"rgb": "FF00FFFF"}); fills.set("count", str(len(fills)))
    for index, case in enumerate(cases, 1):
        formats[index].set("fillId", str(len(fills) - 1))
        if case["mode"] == "none": formats[index].set("borderId", "0")
        elif case["mode"] in ["top", "left"]:
            for edge in list(borders[index]):
                if edge.tag != N + case["mode"]: borders[index].remove(edge)
        elif case["mode"] == "diagonal":
            borders[index].clear(); borders[index].set("diagonalUp", "1")
            edge = ET.SubElement(borders[index], N + "diagonal", {"style": "thin"}); ET.SubElement(edge, N + "color", {"rgb": "FFCC00CC"})
    parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    return native.fixture.zipped(parts)


def fill_points(svg):
    paths = [node for node in ET.fromstring(svg).iter("{http://www.w3.org/2000/svg}path") if node.get("fill") == "rgb(0%, 100%, 100%)"]
    assert len(paths) == 1
    path = paths[0]; values = list(map(float, re.findall(r"[\d.-]+", path.get("d"))))[:8]; assert len(values) == 8
    transform = path.get("transform")
    a, b, c, d, e, f = list(map(float, re.fullmatch(r"matrix\(([^)]+)\)", transform)[1].split(","))) if transform else [1, 0, 0, 1, 0, 0]
    assert b == c == 0
    points = [(values[i] * a + e, values[i + 1] * d + f) for i in range(0, 8, 2)]
    ordered = sorted(points, key=lambda point: (round(point[1], 4), point[0]))
    return [ordered[index] for index in [0, 1, 3, 2]]


def observe(pdf, cases, work):
    native.shared.run(["pdftoppm", "-r", "96", str(pdf), str(work / "raster")])
    rasters = sorted(work.glob("raster-*.ppm")); assert len(rasters) == len(cases)
    for index, case in enumerate(cases, 1):
        svg = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
        points = fill_points(svg.read_bytes())
        if index == 1:
            # Normal fill coverage extends a PDF device unit beyond the actual
            # grid. Use border centerlines, not that raster coverage rectangle;
            # a tiny row-height error is magnified by cot(1 degree).
            vertical, horizontal = set(), set()
            for path in ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"):
                if not re.fullmatch(r"rgb\(79\.99[\d]+%, 0%, 79\.99[\d]+%\)", path.get("stroke", "")): continue
                a, b, c, d, e, f = map(float, re.fullmatch(r"matrix\(([^)]+)\)", path.get("transform"))[1].split(",")); assert b == c == 0
                x1, y1, x2, y2 = map(float, re.findall(r"[\d.-]+", path.get("d")))
                if x1 == x2: vertical.add(x1 * a + e)
                if y1 == y2: horizontal.add(y1 * d + f)
            assert len(vertical) == len(horizontal) == 2
            left, right = sorted(vertical); top, bottom = sorted(horizontal)
        case["cellSizePt"] = [round(right - left, 6), round(bottom - top, 6)]
        case["fillPointsPt"] = [[round(x - left, 6), round(y - top, 6)] for x, y in points]
        with rasters[index - 1].open("rb") as file:
            assert file.readline().strip() == b"P6"; width, height = map(int, file.readline().split()); assert file.readline().strip() == b"255"; pixels = file.read()
        ink = []
        for y in range(int(top / .75) + 1, min(height, int(bottom / .75))):
            for x in range(width):
                red, green, blue = pixels[(y * width + x) * 3:(y * width + x) * 3 + 3]
                # Exclude magenta border strokes; retain black glyph coverage
                # on both white and cyan, including colored antialias pixels.
                if max(red, green, blue) < 224 and not (red > green + 8 and blue > green + 8): ink.append((x, y))
        case["coverageBoundsPt"] = [round(min(x for x, y in ink) * .75 - left, 6), round(min(y for x, y in ink) * .75 - top, 6),
                                    round((max(x for x, y in ink) + 1) * .75 - left, 6), round((max(y for x, y in ink) + 1) * .75 - top, 6)] if ink else None


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    cases = inputs(); data = generate(cases)
    if args.record: FIXTURE.write_bytes(data)
    else: assert FIXTURE.read_bytes() == data
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); path = work / "rotated-cells.xlsx"; path.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(path)])
        (ROOT / ".cache/xlsx-rotated-cells-reference.pdf").write_bytes(path.with_suffix(".pdf").read_bytes())
        observe(path.with_suffix(".pdf"), cases, work)
    report = {"scope": __doc__.strip(), "sourceSha256": sha256(data).hexdigest(), "archiveSha256": native.ARCHIVE_HASH, "environment": environment,
              "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "font": {"family": "Noto Sans CJK SC", "sizePt": 12}, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native rotated-cell observations changed"
    print(f"Verified {len(cases)} native rotated-cell fill/text references")


if __name__ == "__main__": main()
