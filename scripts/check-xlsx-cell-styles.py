"""Observe row/column/cell/default formatting with independent native input."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("native_style_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-cell-style-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-cell-style-precedence.xlsx"
S = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
COLORS = ["#eeeeee", "#ffffdd", "#ddeeff", "#ffdddd", "#ddffdd"]


def generate():
    book = native.fixture.Book(); sheet = book.sheet("Styles"); namespace = "{" + native.fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"])
    formats = styles.find(namespace + "numFmts"); formats.clear(); formats.set("count", "5")
    for index, prefix in enumerate(["N", "D", "C", "R", "X"]):
        ET.SubElement(formats, namespace + "numFmt", {"numFmtId": str(164 + index), "formatCode": '"' + prefix + '"0'})
    fills = styles.find(namespace + "fills")
    # Retain the two reserved fill records, then use distinct complete fills.
    for child in list(fills)[2:]: fills.remove(child)
    for color in COLORS:
        fill = ET.SubElement(fills, namespace + "fill"); pattern = ET.SubElement(fill, namespace + "patternFill", {"patternType": "solid"})
        ET.SubElement(pattern, namespace + "fgColor", {"rgb": "FF" + color[1:].upper()})
    fills.set("count", "7")
    base = {"fontId": "0", "borderId": "0", "applyNumberFormat": "1", "applyFill": "1"}
    named = styles.find(namespace + "cellStyleXfs"); named.clear(); named.set("count", "1")
    ET.SubElement(named, namespace + "xf", {**base, "fillId": "2", "numFmtId": "164"})
    cells = styles.find(namespace + "cellXfs"); cells.clear(); cells.set("count", "4")
    for index in range(4):
        ET.SubElement(cells, namespace + "xf", {**base, "xfId": "0", "fillId": str(index + 3), "numFmtId": str(index + 165)})
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    sheet.before += '<cols><col min="2" max="8" style="1"/></cols>'
    sheet.row_props = {2: {"s": 2, "customFormat": 1}, 3: {"s": 2, "customFormat": 0}, 4: {"s": 2}, 5: {"customFormat": 1}, 6: {"s": 0, "customFormat": 1}}
    for row in range(1, 7):
        for col in ["A", "B", "I"]: sheet.add(col + str(row), 1)
        sheet.add("C" + str(row), 1, style=0); sheet.add("D" + str(row), 1, style=3)
        sheet.add("E" + str(row))  # Stored empty, style absent.
        # F is an implicit blank with no cell record.
        sheet.add("G" + str(row), style=0); sheet.add("H" + str(row), style=3)
    return native.fixture.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        content, styles = [ET.fromstring(archive.read(part)) for part in ["content.xml", "styles.xml"]]
    definitions = {style.get(f"{{{S}}}name"): style for root in [styles, content] for style in root.iter(f"{{{S}}}style")}
    default = styles.find(f".//{{{S}}}default-style[@{{{S}}}family='table-cell']")
    def fill(name, seen=()):
        assert name not in seen
        node = definitions.get(name) if name else default
        if node is None: return None
        properties = node.find(f"{{{S}}}table-cell-properties")
        background = None if properties is None else properties.get(f"{{{FO}}}background-color")
        return background if background is not None else fill(node.get(f"{{{S}}}parent-style-name"), (*seen, name)) if name else None
    table = content.find(f".//{{{native.T}}}table"); columns = []
    for column in table.findall(f"{{{native.T}}}table-column"):
        columns.extend([column.get(f"{{{native.T}}}default-cell-style-name")] * min(9 - len(columns), int(column.get(f"{{{native.T}}}number-columns-repeated", "1"))))
        if len(columns) == 9: break
    samples = []; row_index = 1
    for row in table.findall(f"{{{native.T}}}table-row"):
        column_index = 0
        for cell in row.findall(f"{{{native.T}}}table-cell"):
            for _ in range(min(9 - column_index, int(cell.get(f"{{{native.T}}}number-columns-repeated", "1")))):
                style = cell.get(f"{{{native.T}}}style-name") or row.get(f"{{{native.T}}}default-cell-style-name") or columns[column_index]
                samples.append({"address": chr(65 + column_index) + str(row_index), "text": "".join(cell.itertext()), "fill": fill(style)})
                column_index += 1
            if column_index == 9: break
        assert column_index == 9
        row_index += 1
        if row_index == 7: break
    assert len(samples) == 54
    return samples


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "styles.xlsx"; source.write_bytes(generate())
        if args.record:
            FIXTURE.parent.mkdir(exist_ok=True); FIXTURE.write_bytes(source.read_bytes())
        else: assert FIXTURE.read_bytes() == source.read_bytes(), "Independent style fixture changed"
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        samples = observe(source.with_suffix(".ods"))
        for sample in samples:
            col, row = sample["address"]
            index = 1 if col in "CG" else 4 if col in "DH" else 3 if row == "2" else 1 if row in "56" else 2 if col in "BEF" else 0
            assert sample["fill"] == COLORS[index], sample
            assert sample["text"] == ("" if col in "EFGH" else "NDCRX"[index] + "1"), sample
        report = {"scope": "Native Calc row/column/cell/default formatting including stored and implicit blanks. Normal differs from cellXfs zero; customFormat is tested both absent and false. No Excel/WPS or complete visual acceptance claim.",
                  "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                  "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "samples": samples}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native cell-format precedence changed"
    print(f"Verified {len(samples)} native cell/default format selections")


if __name__ == "__main__": main()
