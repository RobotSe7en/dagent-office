"""Observe independent DrawingML text metrics using pinned Calc and fixed fonts.

This records a producer's run/default/end-paragraph behavior, not SDK reading
acceptance or Excel/WPS behavior. Microsoft documents an 18pt fallback; this
Calc version's fallback and line metrics are recorded separately.
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
spec = importlib.util.spec_from_file_location("native_text_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-drawing-text-semantics.json"
SVG = "{http://www.w3.org/2000/svg}"
HTML = "{http://www.w3.org/1999/xhtml}"


def inputs():
    result = [{"fontPt": size, "endPt": end, "paragraphs": ["Alpha", "Beta"], "insets": "default", "listStyle": "none"}
              for size in [None, 14, 18] for end in [None, 14, 18]]
    result += [{"fontPt": 14, "endPt": end, "paragraphs": ["", "Alpha", "Beta"], "insets": "default", "listStyle": "none"} for end in [None, 14, 18]]
    result += [{"fontPt": 14, "endPt": None, "paragraphs": ["Alpha", "Beta"], "insets": inset, "listStyle": "none"} for inset in ["zero", "explicit-default"]]
    result += [{"fontPt": None, "endPt": None, "paragraphs": ["Alpha", "Beta"], "insets": "default", "listStyle": kind} for kind in ["default-10", "level-22", "both"]]
    result += [{"fontPt": 14, "endPt": end, "breakPt": size, "paragraphs": ["Alpha\n", "Beta"], "insets": "default", "listStyle": "none"}
               for size in [14, 22] for end in [None, 14, 18]]
    result += [{"fontPt": 14, "endPt": 14, "breakPt": size, "paragraphs": paragraphs, "insets": "default", "listStyle": "none"}
               for size in [14, 22] for paragraphs in [["\nAlpha", "Beta"], ["Alpha\n\nBeta"]]]
    return result


def generate(cases):
    book = native.fixture.Book()
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$F$20"; sheet.add("A1", "Case" + str(index))
        shape = native.fixture.basic_shape(1, "").replace('prst="roundRect"', 'prst="rect"').replace("E8F0E6", "0000FF")
        shape = re.sub(r"<a:xfrm[^>]*>.*?</a:xfrm>", '<a:xfrm><a:off x="0" y="0"/><a:ext cx="1524000" cy="1143000"/></a:xfrm>', shape)
        shape = re.sub(r"<a:ln.*?</a:ln>", "<a:ln><a:noFill/></a:ln>", shape)
        paragraphs = ""
        for word in case["paragraphs"]:
            size = "" if case["fontPt"] is None else f' sz="{case["fontPt"] * 100}"'
            end = "" if case["endPt"] is None else f' sz="{case["endPt"] * 100}"'
            break_size = case.get("breakPt", case["fontPt"])
            line_break = '<a:br><a:rPr' + ("" if break_size is None else f' sz="{break_size * 100}"') + '><a:latin typeface="Liberation Sans"/></a:rPr></a:br>'
            run = line_break.join(f'<a:r><a:rPr{size}><a:latin typeface="Liberation Sans"/></a:rPr><a:t>{part}</a:t></a:r>' if part else "" for part in word.split("\n"))
            paragraphs += f'<a:p><a:pPr/>{run}<a:endParaRPr{end}/></a:p>'
        inset = ' lIns="0" rIns="0" tIns="0" bIns="0"' if case["insets"] == "zero" else ' lIns="91440" rIns="91440" tIns="45720" bIns="45720"' if case["insets"] == "explicit-default" else ""
        style = ""
        if case["listStyle"] in ["default-10", "both"]: style += '<a:defPPr><a:defRPr sz="1000"/></a:defPPr>'
        if case["listStyle"] in ["level-22", "both"]: style += '<a:lvl1pPr><a:defRPr sz="2200"/></a:lvl1pPr>'
        body = f'<xdr:txBody><a:bodyPr wrap="square"{inset}/><a:lstStyle>{style}</a:lstStyle>{paragraphs}</xdr:txBody>'
        shape = re.sub(r"<xdr:txBody>.*?</xdr:txBody>", body, shape)
        native.fixture.add_drawing(book, sheet, '<xdr:absoluteAnchor><xdr:pos x="1270000" y="1270000"/><xdr:ext cx="1524000" cy="1143000"/>' + shape + '<xdr:clientData/></xdr:absoluteAnchor>')
    return native.fixture.zipped(book.finish())


def observe(pdf, resaved, cases, work):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//" + HTML + "page")
    assert len(pages) == len(cases)
    results = []
    with ZipFile(resaved) as archive:
        for index, (case, page) in enumerate(zip(cases, pages), 1):
            path = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(path)])
            shapes = [node for node in ET.fromstring(path.read_bytes()).iter(SVG + "path") if node.get("fill") == "rgb(0%, 0%, 100%)"]
            assert len(shapes) == 1 and shapes[0].get("transform") is None
            data = shapes[0].get("d"); assert re.fullmatch(r"[MLZ\d.\s-]+", data)
            points = [(float(x), float(y)) for x, y in re.findall(r"[ML] ([\d.-]+) ([\d.-]+)", data)]
            left = min(p[0] for p in points); top = min(p[1] for p in points)
            width = max(p[0] for p in points) - left; height = max(p[1] for p in points) - top; scale = width / 120
            assert abs(height / scale - 90) < .02
            words = []
            for word in page.findall(HTML + "word"):
                if word.text not in ["Alpha", "Beta"]: continue
                bounds = [float(word.get(key)) for key in ["xMin", "yMin", "xMax", "yMax"]]
                # Font extents may overhang an edge with zero text insets.
                # Record that signed position instead of clipping the evidence.
                assert left - 24 <= bounds[0] < bounds[2] <= left + width + 24 and top - 24 <= bounds[1] < bounds[3] <= top + height + 24, (case, bounds, [left, top, width, height])
                words.append({"text": word.text, "boundsPt": [round((value - (left if axis % 2 == 0 else top)) / scale, 6) for axis, value in enumerate(bounds)]})
            assert [word["text"] for word in words] == ["Alpha", "Beta"]
            source = ET.fromstring(archive.read(f"xl/drawings/drawing{index}.xml")); ns = "{" + native.fixture.A + "}"
            body = source.find(".//" + ns + "bodyPr"); assert body is not None
            text = source.find(".//{" + native.fixture.D + "}txBody"); assert text is not None
            run_sizes = [[node.get("sz") for node in p.iter(ns + "rPr")] for p in text.findall(ns + "p")]
            end_sizes = [None if p.find(ns + "endParaRPr") is None else p.find(ns + "endParaRPr").get("sz") for p in text.findall(ns + "p")]
            results.append({**case, "observed": {"pdfScale": round(scale, 9), "words": words, "lineAdvancePt": round(words[1]["boundsPt"][1] - words[0]["boundsPt"][1], 6),
                "resavedInsets": {name: body.get(name) for name in ["lIns", "rIns", "tIns", "bIns"]}, "resavedRunSizes": run_sizes, "resavedEndSizes": end_sizes}})
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]; assert environment["fonts"] == fonts
    cache = ROOT / ".cache/xlsx-drawing-text"; cache.mkdir(exist_ok=True); cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); source = work / "text.xlsx"; source.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]: native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(source)])
        observations = observe(output / "text.pdf", output / "text.xlsx", cases, work)
    report = {"scope": "Pinned Calc DrawingML text fallback, nonempty/empty paragraph metrics and text insets. Microsoft Office fallback is documented separately; SDK line boxes and Excel/WPS are unverified.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()}, "cases": observations}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native drawing-text evidence changed"
    print(f"Verified {len(cases)} independent drawing-text probes with pinned Calc")


if __name__ == "__main__": main()
