"""Independent native chart color variation and legend observations.

Source ZIP/XML is generated without the SDK. Producer color/marker properties
and PDF word counts are evidence; no font geometry or Excel/WPS run is claimed.
"""
from collections import Counter
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
spec = importlib.util.spec_from_file_location("chart_color_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
C = "{http://schemas.openxmlformats.org/drawingml/2006/chart}"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
HTML = "{http://www.w3.org/1999/xhtml}"
REPORT = ROOT / "docs/evidence/xlsx-chart-colors.json"
FIXTURES = ROOT / "packages/office/tests/fixtures"
ALTERNATE = ["AA5500", "008877", "5544AA", "FF9900", "4488AA", "997700"]


def inputs():
    return ([{"kind": kind, "vary": vary, "series": count} for kind in ["column", "line", "pie"]
             for vary in [False, True] for count in ([1] if kind == "pie" else [1, 2])]
            + [{"kind": "column", "vary": True, "series": 1, "fill": "series"},
               {"kind": "column", "vary": True, "series": 1, "fill": "point"},
               {"kind": "pie", "vary": True, "series": 1, "fill": "point"}])


def generate(cases, alternate=False):
    ET.register_namespace("c", C[1:-1]); ET.register_namespace("a", A[1:-1])
    book = native.fixture.Book()
    if alternate:
        theme = ET.fromstring(book.parts["xl/theme/theme1.xml"])
        for index, color in enumerate(ALTERNATE, 1): theme.find(".//" + A + "accent" + str(index) + "/" + A + "srgbClr").set("val", color)
        book.parts["xl/theme/theme1.xml"] = ET.tostring(theme, encoding="unicode")
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$L$28"
        for row, category in [(40, "One"), (41, "Two"), (42, "Three")]:
            sheet.add("A" + str(row), category); sheet.add("B" + str(row), row - 39)
            if case["series"] == 2: sheet.add("C" + str(row), 43 - row)
        sheet.add("B39", "Series1")
        if case["series"] == 2: sheet.add("C39", "Series2")
        values = [("Series" + str(i + 1), [1, 2, 3] if i == 0 else [3, 2, 1]) for i in range(case["series"])]
        source = native.fixture.drawing.chart_xml(case["kind"], ["One", "Two", "Three"], values)
        source = re.sub(r"<c:externalData\b.*?</c:externalData>", "", source).replace("Data!", sheet.name + "!")
        for old, new in [("$B$1", "$B$39"), ("$C$1", "$C$39"), ("$A$2:$A$4", "$A$40:$A$42"), ("$B$2:$B$4", "$B$40:$B$42"), ("$C$2:$C$4", "$C$40:$C$42")]: source = source.replace(old, new)
        root = ET.fromstring(source); chart = root.find(C + "chart"); plot = chart.find(C + "plotArea")
        graph = plot.find(C + {"column": "barChart", "line": "lineChart", "pie": "pieChart"}[case["kind"]])
        for title in chart.findall(C + "title"): chart.remove(title)
        for series in graph.findall(C + "ser"):
            for style in series.findall(C + "spPr"): series.remove(style)
            if case.get("fill") == "series":
                style = ET.Element(C + "spPr"); solid = ET.SubElement(style, A + "solidFill"); ET.SubElement(solid, A + "srgbClr", {"val": "112233"})
                series.insert(next(i for i, node in enumerate(series) if node.tag == C + "cat"), style)
            if case["kind"] == "line":
                for old in series.findall(C + "marker"): series.remove(old)
                marker = ET.Element(C + "marker"); ET.SubElement(marker, C + "symbol", {"val": "circle"}); ET.SubElement(marker, C + "size", {"val": "7"})
                series.insert(next(i for i, node in enumerate(series) if node.tag == C + "cat"), marker)
            if case.get("fill") == "point":
                point = ET.Element(C + "dPt"); ET.SubElement(point, C + "idx", {"val": "1"}); style = ET.SubElement(point, C + "spPr")
                solid = ET.SubElement(style, A + "solidFill"); ET.SubElement(solid, A + "srgbClr", {"val": "123456"})
                series.insert(next(i for i, node in enumerate(series) if node.tag == C + "cat"), point)
        for flag in graph.findall(C + "varyColors"): graph.remove(flag)
        at = next(i for i, node in enumerate(graph) if node.tag == C + "ser")
        graph.insert(at, ET.Element(C + "varyColors", {"val": str(int(case["vary"]))}))
        path = f"xl/charts/chart{index}.xml"; book.part(path, ET.tostring(root, encoding="unicode"), "chart")
        drawing = native.fixture.add_drawing(book, sheet, native.fixture.one_anchor(native.fixture.graphic_frame(1, "chart"), 0, 0, 480, 300))
        book.relate(drawing, "chart", native.fixture.R + "/chart", path)
    return native.fixture.zipped(book.finish())


def observe(pdf, resaved, cases):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//" + HTML + "page")
    assert len(pages) == len(cases)
    def val(node): return None if node is None else node.get("val")
    def colors(node): return [color.attrib for color in node.findall(".//" + A + "srgbClr")]
    observations = []
    with ZipFile(resaved) as archive:
        for index, (case, page) in enumerate(zip(cases, pages), 1):
            root = ET.fromstring(archive.read(f"xl/charts/chart{index}.xml"))
            words = Counter(word.text for word in page.findall(HTML + "word") if word.text in ["One", "Two", "Three", "Series1", "Series2"])
            series = []
            for item in root.findall(".//" + C + "ser"):
                points = [{"index": val(point.find(C + "idx")), "colors": colors(point), "marker": val(point.find(C + "marker/" + C + "symbol"))} for point in item.findall(C + "dPt")]
                series.append({"index": val(item.find(C + "idx")), "fill": val(item.find(C + "spPr/" + A + "solidFill/" + A + "srgbClr")), "points": points})
            svg = pdf.parent / f"colors-{index}.svg"
            native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(pdf), str(svg)])
            fills = Counter()
            for node in ET.fromstring(svg.read_bytes()).iter():
                if node.tag.rsplit("}", 1)[-1] not in ["path", "rect", "circle", "ellipse", "polygon"]: continue
                color = node.get("fill", "")
                if not color:
                    match = re.search(r"(?:^|;)\s*fill\s*:\s*([^;]+)", node.get("style", "")); color = match.group(1) if match else ""
                match = re.fullmatch(r"rgb\(([^)]+)\)", color)
                if not match: continue
                channels = [round(float(value.strip().rstrip("%")) * (2.55 if "%" in value else 1)) for value in match.group(1).split(",")]
                assert len(channels) == 3 and all(0 <= value <= 255 for value in channels)
                hex_color = "".join(f"{value:02x}" for value in channels)
                if hex_color not in ["000000", "ffffff"]: fills[hex_color] += 1
            observations.append({**case, "producer": {"vary": [flag.attrib for flag in root.findall(".//" + C + "varyColors")], "series": series,
                                                       "pdfWords": dict(sorted(words.items())), "pdfFillColors": sorted(fills)}})
    return observations


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    cache = ROOT / ".cache/xlsx-chart-colors"; cache.mkdir(exist_ok=True); cases = inputs(); results = []
    for name, alternate in [("default", False), ("alternate", True)]:
        data = generate(cases, alternate); fixture = FIXTURES / f"xlsx-chart-colors-{name}.xlsx"
        with TemporaryDirectory(dir=cache) as temporary:
            work = Path(temporary); output = work / "out"; output.mkdir(); source = work / "colors.xlsx"; source.write_bytes(data)
            profile = "-env:UserInstallation=" + (work / "profile").as_uri()
            for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]: native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(source)])
            observed = observe(output / "colors.pdf", output / "colors.xlsx", cases)
        results.append({"theme": name, "sourceSha256": sha256(data).hexdigest(), "cases": observed})
        if args.record: fixture.write_bytes(data)
        else: assert fixture.read_bytes() == data, "Independent chart color source changed"
    report = {"scope": __doc__, "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "workbooks": results,
              "limits": ["No font geometry or pixel equivalence is claimed by these structural colors, PDF vector fill colors and words.", "PDF fill colors are normalized as sets, independent of a PDF-to-SVG tool's subdivision of one colored shape into paths.", "Calc varied lines add per-point markers outside the current explicit point-override reading profile; producer reopen is not claimed here.", "Calc varied-line marker colors and explicit-series-color legend keys differ from the source-themed point colors; these cases are not color-fidelity acceptance.", "Excel/WPS unverified."]}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Pinned chart color observations changed"
    print(f"Verified {len(cases) * len(results)} independent chart color and legend observations")


if __name__ == "__main__": main()
