"""Record fixed-font rotated and stacked native text, independently of the SDK."""
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
spec = importlib.util.spec_from_file_location("rotation_source", ROOT / "scripts/check-xlsx-cell-appearance.py")
source = importlib.util.module_from_spec(spec); spec.loader.exec_module(source)
native = source.native
REPORT = ROOT / "docs/evidence/xlsx-rotation-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-rotation.xlsx"


def inputs():
    result = [{"kind": "rotation", "rotation": rotation, "text": "预算 Ab", "border": rotation == 0} for rotation in [0, 1, 45, 89, 90, 91, 135, 179, 180, 255]]
    result += [{"kind": "rotation", "rotation": 255, "text": text, "border": False} for text in ["ABC", "预算\nAb", "禰󠄀Á😀"]]
    return result


def observe(pdf, cases, work):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//{http://www.w3.org/1999/xhtml}page")
    assert len(pages) == len(cases)
    native.shared.run(["pdftoppm", "-r", "96", str(pdf), str(work / "raster")])
    rasters = sorted(work.glob("raster-*.ppm")); assert len(rasters) == len(cases)
    for index, (page, case) in enumerate(zip(pages, cases), 1):
        svg = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
        vertical, horizontal = set(), set()
        paths = list(ET.fromstring(svg.read_bytes()).iter("{http://www.w3.org/2000/svg}path"))
        for path in paths:
            color = re.fullmatch(r"rgb\(([\d.]+)%, ([\d.]+)%, ([\d.]+)%\)", path.get("stroke", ""))
            if not color or any(abs(float(value) - expected) > .01 for value, expected in zip(color.groups(), [80, 0, 80])): continue
            a, b, c, d, e, f = [float(value) for value in re.fullmatch(r"matrix\(([^)]+)\)", path.get("transform"))[1].split(",")]; assert b == c == 0
            points = re.fullmatch(r"M ([\d.-]+) ([\d.-]+) L ([\d.-]+) ([\d.-]+)\s*", path.get("d")); assert points
            x1, y1, x2, y2 = map(float, points.groups())
            if x1 == x2: vertical.add(round(x1 * a + e, 6))
            elif y1 == y2: horizontal.add(round(y1 * d + f, 6))
        if case["border"]:
            assert len(vertical) == len(horizontal) == 2
            left, right = sorted(vertical); top, bottom = sorted(horizontal); cell_bounds = (left, right, top, bottom)
        else:
            assert not vertical and not horizontal
            left, right, top, bottom = cell_bounds
        words = []
        for word in page.findall("{http://www.w3.org/1999/xhtml}word"):
            if (word.text or "").startswith("Case"): continue
            bounds = [float(word.get(key)) for key in ["xMin", "yMin", "xMax", "yMax"]]
            words.append({"text": word.text, "boundsPt": [round(bounds[0] - left, 6), round(bounds[1] - top, 6), round(bounds[2] - left, 6), round(bounds[3] - top, 6)]})
        case["cellSizePt"] = [round(right - left, 6), round(bottom - top, 6)]; case["pdfWords"] = words
        with rasters[index - 1].open("rb") as file:
            assert file.readline().strip() == b"P6"; width, height = map(int, file.readline().split()); assert file.readline().strip() == b"255"; pixels = file.read()
        def visible(x, y, threshold):
            rgb = pixels[(y * width + x) * 3:(y * width + x) * 3 + 3]
            return max(rgb) < threshold and max(rgb) - min(rgb) < 32
        for name, threshold in [("inkBoundsPt", 64), ("coverageBoundsPt", 224)]:
            ink = [(x, y) for y in range(int(top / .75) + 1, min(height, int(bottom / .75))) for x in range(width) if visible(x, y, threshold)]
            assert ink
            case[name] = [round(min(x for x, y in ink) * .75 - left, 6), round(min(y for x, y in ink) * .75 - top, 6),
                          round((max(x for x, y in ink) + 1) * .75 - left, 6), round((max(y for x, y in ink) + 1) * .75 - top, 6)]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = environment.pop("fonts"); environment["fontManifestSha256"] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest(); environment["fontCount"] = len(fonts)
    cases = inputs(); data = source.generate(cases, 180)
    if args.record: FIXTURE.write_bytes(data)
    else: assert FIXTURE.read_bytes() == data
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); path = work / "rotation.xlsx"; path.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(path)])
        (ROOT / ".cache/xlsx-rotation-reference.pdf").write_bytes(path.with_suffix(".pdf").read_bytes())
        observe(path.with_suffix(".pdf"), cases, work)
    report = {"scope": "Native fixed-font PDF positions for borderless rotated and stacked text. Cell bounds come from the identical unrotated first sheet. Rotated-border geometry is separate. Producer Unicode splitting is recorded, not adopted as SDK behavior. Not Excel/WPS acceptance.",
              "sourceSha256": sha256(data).hexdigest(), "archiveSha256": native.ARCHIVE_HASH, "environment": environment,
              "pixelThresholds": {"darkInk": 64, "visibleCoverage": 224, "maximumChannelSpread": 32},
              "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "font": {"family": "Noto Sans CJK SC", "sizePt": 12}, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native rotation observations changed"
    print(f"Verified {len(cases)} native rotated and stacked text references")


if __name__ == "__main__": main()
