"""Observe color-scale boundary paints in fixed Calc PDF, independently of SDK."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("scale_native_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-color-scale-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-color-scales.xlsx"
SVG = "{http://www.w3.org/2000/svg}"
CASES = [("two_constant", [5, 5, 5], 2), ("two_zero", [0, 0, 0], 2), ("two_mixed", [-100, 0, 100], 2),
         ("three_constant", [5, 5, 5], 3), ("three_zero", [0, 0, 0], 3), ("three_mixed", [-100, 0, 100], 3),
         ("three_repeated_min", [0, 0, 100], 3), ("three_repeated_max", [0, 100, 100], 3)]


def generate():
    book = native.fixture.Book()
    for name, values, count in CASES:
        sheet = book.sheet(name); sheet.print_area = "$A$1:$A$3"
        for index, value in enumerate(values, 1): sheet.add("A" + str(index), value); sheet.row_props[index] = {"ht": 30, "customHeight": 1}
        stops = '<cfvo type="min"/>' + ('<cfvo type="percentile" val="50"/>' if count == 3 else '') + '<cfvo type="max"/>'
        colors = '<color rgb="FFFF0000"/>' + ('<color rgb="FFFFFF00"/>' if count == 3 else '') + '<color rgb="FF00FF00"/>'
        sheet.after += f'<conditionalFormatting sqref="A1:A3"><cfRule type="colorScale" priority="1"><colorScale>{stops}{colors}</colorScale></cfRule></conditionalFormatting>'
    return native.fixture.zipped(book.finish())


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    data = generate(); cache = ROOT / ".cache/xlsx-color-scales"; cache.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); path = work / "scales.xlsx"; path.write_bytes(data)
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "pdf", "--outdir", str(work), str(path)])
        records = []
        for index, (name, values, count) in enumerate(CASES, 1):
            svg = work / f"page-{index}.svg"; native.shared.run(["pdftocairo", "-f", str(index), "-l", str(index), "-svg", str(path.with_suffix('.pdf')), str(svg)])
            fills = [node.get("fill") for node in ET.fromstring(svg.read_bytes()).iter(SVG + "path")
                     if node.get("fill", "").startswith("rgb(") and node.get("fill") not in ["rgb(0%, 0%, 0%)", "rgb(100%, 100%, 100%)"]]
            records.append({"name": name, "values": values, "stops": count, "pdfPathFills": fills})
    report = {"scope": "Pinned Calc PDF path paints for constant and repeated color-scale thresholds. Not Excel/WPS or SDK acceptance.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(), "cases": records}
    if args.record: FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == data, "Independent color-scale source changed"
        assert json.loads(REPORT.read_text()) == report, "Native color-scale paint changed"
    print(json.dumps(records, indent=2))


if __name__ == "__main__": main()
