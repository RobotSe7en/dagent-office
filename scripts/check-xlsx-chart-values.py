"""Record independent chart cache and plotted-data behavior with pinned Calc.

This is producer evidence, not an Excel/WPS claim. Deliberately stale chart
caches differ from worksheet source cells; no SDK importer or exporter is used.
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
spec = importlib.util.spec_from_file_location("chart_value_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-chart-value-semantics.json"
C = "{" + native.fixture.C + "}"
SVG = "{http://www.w3.org/2000/svg}"
HTML = "{http://www.w3.org/1999/xhtml}"


def inputs():
    cases = [{"kind": kind, "blankAs": blanks, "hidden": "none", "visibleOnly": True} for kind in ["column", "line", "pie"] for blanks in ["gap", "zero", "span"]]
    cases += [{"kind": "column", "blankAs": "gap", "hidden": hidden, "visibleOnly": visible} for hidden in ["rows", "values-column", "categories-column"] for visible in [False, True]]
    return cases


def generate(cases):
    book = native.fixture.Book()
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$J$30"
        sheet.add("A1", "Category"); sheet.add("B1", "Values")
        for row in range(2, 14): sheet.add("A" + str(row), "Point" + str(row - 2))
        for row, value in [(2, 2), (3, "7"), (4, "word"), (5, True), (6, False), (12, -3), (13, 0)]: sheet.add("B" + str(row), value)
        sheet.add("B8", style=0); sheet.add("B9", "", kind="str", formula='""')
        sheet.add("B10", "#N/A", kind="e"); sheet.add("B11", "#DIV/0!", kind="e")
        if case["hidden"] == "rows": sheet.row_props = {3: {"hidden": 1}, 5: {"hidden": 1}, 12: {"hidden": 1}}
        if case["hidden"].endswith("-column"):
            column = 2 if case["hidden"] == "values-column" else 1
            sheet.before += f'<cols><col min="{column}" max="{column}" hidden="1" width="12" customWidth="1"/></cols>'
        frame = native.fixture.graphic_frame(1, "chart")
        drawing = native.fixture.add_drawing(book, sheet, native.fixture.one_anchor(frame, 2, 2, 360, 240))
        source = native.fixture.drawing.chart_xml(case["kind"], ["Cache" + str(i) for i in range(12)], [("CachedSeries", [999] * 12)])
        source = re.sub(r'<c:externalData\b.*?</c:externalData>', '', source).replace("Data!", sheet.name + "!")
        source = source.replace('<c:dispBlanksAs val="gap"/>', f'<c:dispBlanksAs val="{case["blankAs"]}"/>').replace('<c:plotVisOnly val="1"/>', f'<c:plotVisOnly val="{int(case["visibleOnly"])}"/>')
        path = f"xl/charts/chart{index}.xml"; book.part(path, source, "chart"); book.relate(drawing, "chart", native.fixture.R + "/chart", path)
    return native.fixture.zipped(book.finish())


def blue(value):
    if not value: return False
    match = re.fullmatch(r"rgb\(([\d.]+)%,\s*([\d.]+)%,\s*([\d.]+)%\)", value)
    return bool(match and all(abs(float(actual) * 2.55 - expected) < .1 for actual, expected in zip(match.groups(), [68, 114, 196])))


def observe(pdf, resaved, cases, work):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//" + HTML + "page")
    assert len(pages) == len(cases), len(pages)
    observations = []
    with ZipFile(resaved) as archive:
        for index, (case, page) in enumerate(zip(cases, pages), 1):
            chart = ET.fromstring(archive.read(f"xl/charts/chart{index}.xml")); series = chart.findall(".//" + C + "ser"); assert len(series) == 1
            references = []
            for role in ["tx", "cat", "val"]:
                reference = series[0].find(C + role); assert reference is not None
                vector = list(reference)[0]; formula = vector.find(C + "f"); cache = next((child for child in vector if child.tag in [C + "strCache", C + "numCache"]), None)
                assert formula is not None and cache is not None
                count = cache.find(C + "ptCount"); assert count is not None
                references.append({"role": role, "kind": vector.tag.removeprefix(C), "formula": formula.text, "count": int(count.get("val")),
                                   "points": [{"index": int(point.get("idx")), "value": point.find(C + "v").text} for point in cache.findall(C + "pt")]})
            svg = work / f"chart-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
            paths = []
            def geometry(node, transforms):
                # Native negative bars may reflect a positive local path in an
                # ancestor group. Retain that transform instead of treating
                # the local coordinates as final page positions.
                local = transforms + ([node.get("transform")] if node.get("transform") else [])
                if node.tag == SVG + "path" and (blue(node.get("fill")) or blue(node.get("stroke"))):
                    paths.append({"attributes": dict(node.attrib), "transforms": local})
                for child in node: geometry(child, local)
            geometry(ET.fromstring(svg.read_bytes()), [])
            words = [word.text for word in page.findall(HTML + "word")]
            observations.append({**case, "observed": {"references": references, "blueGeometry": paths, "pdfWords": words}})
    return observations


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]; assert environment["fonts"] == fonts
    cache = ROOT / ".cache/xlsx-chart-values"; cache.mkdir(exist_ok=True); cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); source = work / "chart-values.xlsx"; source.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]: native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(source)])
        observations = observe(output / "chart-values.pdf", output / "chart-values.xlsx", cases, work)
    report = {"scope": "Independent chart source/cache mismatch, mixed cell types, blanks and hidden data. Calc producer observations only; no Excel/WPS or SDK rendering acceptance.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()}, "cases": observations}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native chart-value evidence changed"
    print(f"Verified {len(cases)} independent chart-value cases with pinned Calc")


if __name__ == "__main__": main()
