"""Observe unblurred outer-shadow direction using independent XLSX and pinned Calc.

The recorded producer ignores rotWithShape in these cases. This is not evidence
that Excel/WPS share that behavior, nor a complete DrawingML appearance gate.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import math
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("native_effect_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-drawing-shadow-semantics.json"
SVG = "{http://www.w3.org/2000/svg}"


def inputs():
    return [{"rotation": angle, "flipH": flip_h, "flipV": flip_v, "rotateWithShape": rotate, "direction": direction, "distancePt": 15}
            for angle in [0, 30] for flip_h, flip_v in [(False, False), (True, False), (False, True), (True, True)]
            for rotate in [False, True] for direction in [0, 90]]


def generate(cases):
    book = native.fixture.Book()
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$F$20"; sheet.add("A1", "Case" + str(index))
        shape = native.fixture.basic_shape(1, "")
        transform = f'<a:xfrm rot="{case["rotation"] * 60000}" flipH="{int(case["flipH"])}" flipV="{int(case["flipV"])}"><a:off x="0" y="0"/><a:ext cx="1524000" cy="762000"/></a:xfrm>'
        shape = re.sub(r"<a:xfrm[^>]*>.*?</a:xfrm>", transform, shape)
        shape = re.sub(r"<a:ln.*?</a:ln>", "<a:ln><a:noFill/></a:ln>", shape).replace("E8F0E6", "0000FF")
        shadow = f'<a:effectLst><a:outerShdw blurRad="0" dist="190500" dir="{case["direction"] * 60000}" rotWithShape="{int(case["rotateWithShape"])}"><a:srgbClr val="FF0000"/></a:outerShdw></a:effectLst>'
        shape = shape.replace("</xdr:spPr>", shadow + "</xdr:spPr>")
        native.fixture.add_drawing(book, sheet, '<xdr:absoluteAnchor><xdr:pos x="1270000" y="1270000"/><xdr:ext cx="1524000" cy="762000"/>' + shape + '<xdr:clientData/></xdr:absoluteAnchor>')
    return native.fixture.zipped(book.finish())


def observe(pdf, resaved, cases, work):
    results = []
    with ZipFile(resaved) as archive:
        for index, case in enumerate(cases, 1):
            path = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(path)])
            colors = {}
            for node in ET.fromstring(path.read_bytes()).iter(SVG + "path"):
                if node.get("fill") not in ["rgb(100%, 0%, 0%)", "rgb(0%, 0%, 100%)"]: continue
                assert node.get("transform") is None and node.get("fill") not in colors
                assert re.fullmatch(r"[MLZ\d.\s-]+", node.get("d"))
                colors[node.get("fill")] = [(float(x), float(y)) for x, y in re.findall(r"[ML] ([\d.-]+) ([\d.-]+)", node.get("d"))]
            body, shadow = colors["rgb(0%, 0%, 100%)"], colors["rgb(100%, 0%, 0%)"]
            assert len(body) == len(shadow) and len(body) >= 4
            scale = max(math.dist(a, b) for a, b in zip(body, body[1:])) / 120
            offsets = [((b[0] - a[0]) / scale, (b[1] - a[1]) / scale) for a, b in zip(body, shadow)]
            offset = [sum(point[axis] for point in offsets) / len(offsets) for axis in [0, 1]]
            expected = [15 if case["direction"] == 0 else 0, 15 if case["direction"] == 90 else 0]
            assert all(abs(offset[axis] - expected[axis]) < .02 for axis in [0, 1]), (case, offset)
            assert all(abs(point[axis] - offset[axis]) < .02 for point in offsets for axis in [0, 1])
            source = ET.fromstring(archive.read(f"xl/drawings/drawing{index}.xml"))
            effect = source.find(".//{" + native.fixture.A + "}outerShdw"); assert effect is not None
            results.append({**case, "observed": {"offsetPt": [round(value, 6) for value in offset], "pdfScale": round(scale, 9), "resavedRotateWithShape": effect.get("rotWithShape")}})
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]; assert environment["fonts"] == fonts
    cache = ROOT / ".cache/xlsx-drawing-effects"; cache.mkdir(parents=True, exist_ok=True); cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); source = work / "shadows.xlsx"; source.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]:
            native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(source)])
        observations = observe(output / "shadows.pdf", output / "shadows.xlsx", cases, work)
    report = {"scope": "Unblurred outer-shadow offsets only. Pinned Calc ignores the source rotate-with-shape flag in these cases; Excel/WPS and SDK effects are unverified.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()}, "cases": observations}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native drawing-shadow evidence changed"
    print(f"Verified {len(cases)} independent outer-shadow direction probes with pinned Calc")


if __name__ == "__main__": main()
