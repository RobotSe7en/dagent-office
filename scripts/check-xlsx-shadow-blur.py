"""Record independent native outer-shadow blur profiles with pinned Calc.

The samples describe this producer's rendering, not an assumed Gaussian mapping
or an SDK/Excel/WPS acceptance claim. Source and font/environment hashes are fixed.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from io import BytesIO
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("shadow_directions", ROOT / "scripts/check-xlsx-drawing-effects.py")
directions = importlib.util.module_from_spec(spec); spec.loader.exec_module(directions)
native = directions.native
REPORT = ROOT / "docs/evidence/xlsx-shadow-blur-semantics.json"
SVG = "{http://www.w3.org/2000/svg}"


def inputs():
    return [{"rotation": 0, "flipH": False, "flipV": False, "rotateWithShape": False, "direction": 0, "distancePt": 60, "blurPt": blur} for blur in [0, 4, 12, 24]]


def generate(cases):
    with ZipFile(BytesIO(directions.generate(cases))) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f"xl/drawings/drawing{index}.xml"
        assert parts[path].count(b'blurRad="0" dist="190500"') == 1
        parts[path] = parts[path].replace(b'blurRad="0" dist="190500"', f'blurRad="{case["blurPt"] * 12700}" dist="{case["distancePt"] * 12700}"'.encode())
    return native.fixture.zipped(parts)


def observe(pdf, resaved, cases, work):
    results = []
    with ZipFile(resaved) as archive:
        for index, case in enumerate(cases, 1):
            svg = work / f"page-{index}.svg"; prefix = work / f"raster-{index}"
            native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
            blue = [node for node in ET.fromstring(svg.read_bytes()).iter(SVG + "path") if node.get("fill") == "rgb(0%, 0%, 100%)"]
            assert len(blue) == 1 and blue[0].get("transform") is None
            path = blue[0].get("d"); assert re.fullmatch(r"[MLZ\d.\s-]+", path)
            points = [(float(x), float(y)) for x, y in re.findall(r"[ML] ([\d.-]+) ([\d.-]+)", path)]
            left = min(p[0] for p in points); right = max(p[0] for p in points); top = min(p[1] for p in points); bottom = max(p[1] for p in points)
            scale = (right - left) / 120; assert abs((bottom - top) / scale - 60) < .02
            native.shared.run(["pdftoppm", "-f", str(index), "-l", str(index), "-singlefile", "-r", "144", str(pdf), str(prefix)])
            data = prefix.with_suffix(".ppm").read_bytes(); header = re.match(rb"P6\s+(\d+)\s+(\d+)\s+255\s", data); assert header
            width, height = map(int, header.groups()); pixels = data[header.end():]; assert len(pixels) == width * height * 3
            y = int((top + bottom)); edge = right + case["distancePt"] * scale; samples = []
            edge_calibration = {}
            if not case["blurPt"]:
                row = [(x, pixels[(y * width + x) * 3 + 1]) for x in range(int(edge * 2) - 6, int(edge * 2) + 7)]
                crossings = [x + .5 + (127.5 - a) / (b - a) for (x, a), (_, b) in zip(row, row[1:]) if a < 127.5 <= b]
                assert len(crossings) == 1
                offset = (crossings[0] / 2 - edge) / scale; assert abs(offset) < 1
                edge_calibration["sharpEdgeOffsetPt"] = round(offset, 6)
            for offset in [-24, -18, -12, -8, -6, -4, -2, 0, 2, 4, 6, 8, 12, 18, 24]:
                x = round((edge + offset * scale) * 2 - .5); assert 0 <= x < width and 0 <= y < height
                start = (y * width + x) * 3
                samples.append({"offsetPt": round(((x + .5) / 2 - edge) / scale, 6), "rgb": list(pixels[start:start + 3])})
            drawing = ET.fromstring(archive.read(f"xl/drawings/drawing{index}.xml"))
            shadow = drawing.find(".//{" + native.fixture.A + "}outerShdw"); assert shadow is not None
            results.append({**case, "observed": {"pdfScale": round(scale, 9), "sampleYRelativePt": round(((y + .5) / 2 - top) / scale, 6), "resavedBlurRad": shadow.get("blurRad"), **edge_calibration, "samples": samples}})
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]; assert environment["fonts"] == fonts
    cache = ROOT / ".cache/xlsx-shadow-blur"; cache.mkdir(exist_ok=True); cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); source = work / "shadow-blur.xlsx"; source.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]: native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(source)])
        observations = observe(output / "shadow-blur.pdf", output / "shadow-blur.xlsx", cases, work)
    report = {"scope": "Four independent outer-shadow blur profiles at 144 dpi, normalized to source points. This is producer evidence; no SDK Gaussian mapping or Excel/WPS acceptance is implied.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()}, "cases": observations}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native shadow-blur evidence changed"
    print(f"Verified {len(cases)} independent shadow-blur profiles with pinned Calc")


if __name__ == "__main__": main()
