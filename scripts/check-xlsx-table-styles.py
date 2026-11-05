"""Observe table styles versus cell formatting and hidden rows in pinned Calc.

Independent ZIP/XML inputs; records producer behavior, not Excel/WPS acceptance.
"""
from copy import deepcopy
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
spec = importlib.util.spec_from_file_location("table_style_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-table-style-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-table-style-precedence.xlsx"
S = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
CASES = [("implicit", None), ("explicit_zero", 0), ("bold", 1), ("red_font", 2), ("family", 3), ("duplicate_font", 4),
         ("red_fill", 5), ("duplicate_none_fill", 6), ("duplicate_empty_border", 7), ("bottom_border", 8), ("number_format", 9),
         ("hidden_rows", None), ("hidden_column", None), ("both_stripes", None), ("no_header", None), ("no_totals", None), ("first_last", None)]


def generate():
    book = native.fixture.Book(); ns = "{" + native.fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"])
    fonts, fills, borders = [styles.find(ns + tag) for tag in ["fonts", "fills", "borders"]]
    font_ids = []
    for mode in ["bold", "red", "family", "copy"]:
        font = deepcopy(fonts[0]); font_ids.append(len(fonts))
        if mode == "bold": ET.SubElement(font, ns + "b")
        if mode == "red":
            for node in list(font):
                if node.tag == ns + "color": font.remove(node)
            ET.SubElement(font, ns + "color", {"rgb": "FFCC0000"})
        if mode == "family":
            font.find(ns + "name").set("val", "DejaVu Sans")
            for node in list(font):
                if node.tag == ns + "scheme": font.remove(node)
        fonts.append(font)
    red_fill = len(fills); fill = ET.SubElement(fills, ns + "fill"); pattern = ET.SubElement(fill, ns + "patternFill", {"patternType": "solid"})
    ET.SubElement(pattern, ns + "fgColor", {"rgb": "FFFF0000"})
    empty_fill = len(fills); fills.append(deepcopy(fills[0]))
    empty_border = len(borders); borders.append(deepcopy(borders[0]))
    bottom_border = len(borders); border = ET.SubElement(borders, ns + "border"); bottom = ET.SubElement(border, ns + "bottom", {"style": "thin"}); ET.SubElement(bottom, ns + "color", {"rgb": "FF00FF00"})
    for group in [fonts, fills, borders]: group.set("count", str(len(group)))
    xfs = styles.find(ns + "cellXfs"); xfs.clear(); xfs.set("count", "10")
    base = {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}
    for variant in [{}, *[{"fontId": str(index)} for index in font_ids], {"fillId": str(red_fill)}, {"fillId": str(empty_fill)},
                    {"borderId": str(empty_border)}, {"borderId": str(bottom_border)}, {"numFmtId": "2"}]:
        ET.SubElement(xfs, ns + "xf", {**base, **variant})
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    for index, (name, style) in enumerate(CASES, 1):
        sheet = book.sheet(name); header = name != "no_header"; totals = name != "no_totals"
        if name == "hidden_rows": sheet.row_props = {3: {"hidden": 1}, 5: {"hidden": 1}}
        if name == "hidden_column": sheet.before += '<cols><col min="2" max="2" hidden="1" width="12"/></cols>'
        for row in range(1, 9):
            for col in ["A", "B", "C"]:
                options = {} if style is None else {"style": style}
                sheet.add(col + str(row), col if row == 1 else row, **options)
        columns = ''.join(f'<tableColumn id="{column}" name="{name}"/>' for column, name in enumerate(["A", "B", "C"], 1))
        both = int(name == "both_stripes"); ends = int(name == "first_last")
        autofilter = f'<autoFilter ref="A1:C{7 if totals else 8}"/>' if header else ''
        table = f'<table xmlns="{native.fixture.S}" id="{index}" name="Probe{index}" displayName="Probe{index}" ref="A1:C8" headerRowCount="{int(header)}" totalsRowCount="{int(totals)}">{autofilter}<tableColumns count="3">{columns}</tableColumns><tableStyleInfo name="TableStyleMedium9" showFirstColumn="{ends}" showLastColumn="{ends}" showRowStripes="1" showColumnStripes="{both}"/></table>'
        path = f"xl/tables/table{index}.xml"; book.part(path, native.fixture.XML + table, "table"); book.relate(sheet.path, "table", native.fixture.R + "/table", path)
        sheet.tail = '<tableParts count="1"><tablePart r:id="table"/></tableParts>'
    return native.fixture.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive:
        content, styles = [ET.fromstring(archive.read(part)) for part in ["content.xml", "styles.xml"]]
    definitions = {style.get(f"{{{S}}}name"): style for root in [styles, content] for style in root.iter(f"{{{S}}}style")}
    default = styles.find(f".//{{{S}}}default-style[@{{{S}}}family='table-cell']")
    def appearance(name, seen=()):
        assert name not in seen
        node = definitions.get(name) if name else default
        if node is None: return {}
        result = appearance(node.get(f"{{{S}}}parent-style-name"), (*seen, name)) if name else {}
        for tag in ["table-cell-properties", "text-properties"]:
            properties = node.find(f"{{{S}}}" + tag)
            if properties is not None:
                result.update({key.split('}')[-1]: value for key, value in properties.attrib.items() if key.startswith("{" + FO + "}") or key == "{" + S + "}font-name"})
        return result
    results = []
    for table in content.findall(f".//{{{native.T}}}table"):
        name = table.get(f"{{{native.T}}}name"); columns = []
        for column in table.findall(f"{{{native.T}}}table-column"):
            columns.extend([column.get(f"{{{native.T}}}default-cell-style-name")] * min(3 - len(columns), int(column.get(f"{{{native.T}}}number-columns-repeated", "1"))))
            if len(columns) == 3: break
        samples = []; row_index = 1
        for row in table.iter(f"{{{native.T}}}table-row"):
            for _ in range(min(8 - row_index + 1, int(row.get(f"{{{native.T}}}number-rows-repeated", "1")))):
                column_index = 0
                for cell in row.findall(f"{{{native.T}}}table-cell"):
                    for _ in range(min(3 - column_index, int(cell.get(f"{{{native.T}}}number-columns-repeated", "1")))):
                        style = cell.get(f"{{{native.T}}}style-name") or row.get(f"{{{native.T}}}default-cell-style-name") or columns[column_index]
                        values = appearance(style)
                        samples.append({"address": chr(65 + column_index) + str(row_index), "text": ''.join(cell.itertext()),
                            "visibility": row.get(f"{{{native.T}}}visibility", "visible"),
                            "appearance": {key: values.get(key) for key in ["background-color", "color", "font-weight", "font-name", "border", "border-bottom", "border-top"]}})
                        column_index += 1
                    if column_index == 3: break
                assert column_index == 3
                row_index += 1
            if row_index > 8: break
        assert len(samples) == 24
        results.append({"case": name, "cells": samples})
    assert len(results) == len(CASES)
    return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "tables.xlsx"; source.write_bytes(generate())
        if args.record: FIXTURE.write_bytes(source.read_bytes())
        else: assert FIXTURE.read_bytes() == source.read_bytes(), "Independent table fixture changed"
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        samples = observe(source.with_suffix('.ods'))
        # The producer does not render the referenced built-in table style in
        # these inputs. Do not infer table/direct-format precedence from an
        # absent table layer, or label this as table-style acceptance.
        by_name = {sample["case"]: sample["cells"] for sample in samples}
        baseline = by_name["implicit"]
        assert all(cell["appearance"]["background-color"] is None for cell in baseline)
        for name in ["explicit_zero", "both_stripes", "no_header", "no_totals", "first_last"]:
            assert by_name[name] == baseline, "Pinned Calc started displaying table regions; reassess the observation"
        assert all(cell["appearance"]["font-weight"] == "bold" for cell in by_name["bold"])
        assert all(cell["appearance"]["color"] == "#cc0000" for cell in by_name["red_font"])
        assert all(cell["appearance"]["background-color"] == "#ff0000" for cell in by_name["red_fill"])
        # Share repeated appearance records without losing any cell samples.
        appearances = []
        for sample in samples:
            for cell in sample["cells"]:
                appearance = cell.pop("appearance")
                if appearance not in appearances: appearances.append(appearance)
                cell["appearance"] = appearances.index(appearance)
        report = {"scope": "Pinned Calc does not display the referenced TableStyleMedium9 in these independently generated inputs. Direct cell formatting is observed, but table precedence and hidden-row banding cannot be verified by this producer. No Excel/WPS or SDK composition acceptance claim.",
                  "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                  "appearances": appearances, "samples": samples}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else: assert json.loads(REPORT.read_text()) == report, "Native table formatting changed"
    print(f"Verified {len(samples)} native table-style cases")


if __name__ == '__main__': main()
