"""Record independent native text overflow flags with pinned Calc.

The SVG ink and resaved attributes describe this producer only. They do not
establish Excel/WPS behavior or replace the normative clip/ellipsis semantics.
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
spec = importlib.util.spec_from_file_location("overflow_text_source", ROOT / "scripts/check-xlsx-drawing-text.py")
source = importlib.util.module_from_spec(spec); spec.loader.exec_module(source)
native = source.native
REPORT = ROOT / "docs/evidence/xlsx-text-overflow.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-text-overflow.xlsx"
HTML = source.HTML


def inputs():
    return [{"fontPt": 14, "endPt": 14, "paragraphs": ["AlphaABCDEFGHIJKLMN", "Beta0123456789", "Gamma", "Delta"],
             "insets": "zero", "listStyle": "none", "horizontal": horizontal, "vertical": vertical, "wrap": wrap, "anchor": anchor}
            for horizontal in ["overflow", "clip"] for vertical in ["overflow", "clip", "ellipsis"]
            for wrap in ["none", "square"] for anchor in ["t", "b"]]


def generate(cases):
    with ZipFile(BytesIO(source.generate(cases))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f"xl/drawings/drawing{index}.xml"; xml = parts[path].decode()
        xml = xml.replace('cy="1143000"', 'cy="508000"')  # 120 x 40pt
        xml = xml.replace('<a:bodyPr wrap="square"', f'<a:bodyPr horzOverflow="{case["horizontal"]}" vertOverflow="{case["vertical"]}" anchor="{case["anchor"]}" wrap="{case["wrap"]}"')
        xml = xml.replace('bIns="0"/>', 'bIns="0"><a:noAutofit/></a:bodyPr>')
        parts[path] = xml
    return native.fixture.zipped(parts)


def observe(pdf, resaved, cases, work):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//" + HTML + "page")
    assert len(pages) == len(cases)
    results = []
    with ZipFile(resaved) as archive:
        for index, (case, page) in enumerate(zip(cases, pages), 1):
            svg = work / f"page-{index}.svg"
            native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
            # Keep the vector observation reproducible without adding a raster
            # dependency to the Office job. Browser tests separately inspect ink.
            root = ET.fromstring(svg.read_bytes())
            uses = [node for node in root.iter(source.SVG + "use") if node.get("{http://www.w3.org/1999/xlink}href")]
            words = [{"text": word.text, "bounds": [round(float(word.get(key)), 6) for key in ["xMin", "yMin", "xMax", "yMax"]]}
                     for word in page.findall(HTML + "word") if word.text != "Case" + str(index)]
            body = ET.fromstring(archive.read(f"xl/drawings/drawing{index}.xml")).find(".//{" + native.fixture.A + "}bodyPr")
            assert body is not None
            results.append({**case, "observed": {"words": words, "paintedGlyphUses": len(uses),
                "resaved": {name: body.get(name) for name in ["horzOverflow", "vertOverflow", "wrap", "anchor"]}}})
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    # Every painted character in these probes (including Case labels and the
    # ellipsis) uses regular Liberation Sans. Pin its actual resolved file;
    # unrelated installed fallback fonts must not block this Latin/ellipsis probe.
    baseline_fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]
    expected_font = next(font for font in baseline_fonts if font["file"] == "LiberationSans-Regular.ttf")
    family, filename = native.shared.run(["fc-match", "-f", "%{family}|%{file}", "Liberation Sans:style=Regular"]).split("|")
    assert "Liberation Sans" in family and Path(filename).name == expected_font["file"], "Required regular Liberation Sans did not resolve"
    assert sha256(Path(filename).read_bytes()).hexdigest() == expected_font["sha256"], "Required reference font bytes differ"
    fonts = [expected_font]
    cache = ROOT / ".cache/xlsx-text-overflow"; cache.mkdir(exist_ok=True)
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); path = work / "overflow.xlsx"; path.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]:
            native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(path)])
        observations = observe(output / "overflow.pdf", output / "overflow.xlsx", cases, work)
    report = {"scope": "Pinned Calc observations of DrawingML horizontal/vertical overflow, wrapping and top/bottom anchors; PDF text extraction can include clipped glyphs. Excel/WPS unverified.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"families": ["Liberation Sans"], "files": fonts, "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()},
              "cases": observations}
    if args.record:
        FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == data, "Independent text overflow source changed"
        assert json.loads(REPORT.read_text()) == report, "Native text overflow evidence changed"
    print(f"Verified {len(cases)} independent text-overflow probes with pinned Calc")


if __name__ == "__main__": main()
