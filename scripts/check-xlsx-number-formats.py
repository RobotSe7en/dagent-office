"""Record independent formatted text from native XLSX using pinned Calc.

This observes the producer, including historical dates that differ from Excel.
It does not declare SDK formatting, column fitting or glyph layout verified.
"""
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
BUILD = "LibreOffice 25.8.1.1 54047653041915e595ad4e45cccea684809c77b5"
ARCHIVE_HASH = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
REPORT = ROOT / "docs/evidence/xlsx-number-format-semantics.json"
T = "urn:oasis:names:tc:opendocument:xmlns:table:1.0"
O = "urn:oasis:names:tc:opendocument:xmlns:office:1.0"
TEXT = "urn:oasis:names:tc:opendocument:xmlns:text:1.0"
CASES = [(0, "yyyy-mm-dd"), (1, "yyyy-mm-dd"), (59, "yyyy-mm-dd"), (60, "yyyy-mm-dd"), (61, "yyyy-mm-dd"),
         (59.9 / 86400, "hh:mm"), (59.9 / 86400, "hh:mm:ss"), (59.9 / 86400, "hh:mm:ss.00"), (.5, "h:mm AM/PM"), (1.25, "[h]:mm"),
         (1.005, "0.00"), (-1.005, "0.00"), (.47, "#.##"), (0, "#.##"), (0, "0.##"), (12345.678, "#,##0.00"),
         (-1234.5, "#,##0.00;[Red](#,##0.00)"), (0, '0.00;[Red](0.00);"—";@'), (12200000, '0.0,,"M"'),
         (1234, "00.0E+0"), (1234, "##0.0E+0"), (1234, "0.00E+00"), (1.25, "# ?/?"), (1.3, "# ??/??"),
         (.5, "0.00%%"), (-1, "[>=100]0;[<0](0);0"), (123456789012345, "General"),
         (0, "0#"), (0, "0?"), (0, "?0#"), (0, "00?"), (0, "?#0"), (0, "0.?0"), (1, "0.?0"), (0, ".00"), (0, "0.#0"), (1, "0.#0")]
COMMON = ["00:00", "00:00:59", "00:00:59.90", "12:00 PM", "30:00", "1.01", "-1.01", ".47", "", "0", "12,345.68", "(1,234.50)", "—", "12.2M", "12.3E+2", "1.2E+3", "1.23E+03", "1 1/4", "1 \u20073/10", "50.00%%", "(1)", "123456789012345"]
COMMON += ["0", "0\u2007", "\u20070", "00\u2007", "\u20070", "0.\u20070", "1.\u20070", ".00", "0.00", "1.00"]
CASES += [(value, code) for code in ["# ?/?", "# ??/??", "# ?/8", "?/?", "0/0"] for value in [0, 1, 1.3, 1.99, .01]]
COMMON += ["0 \u2007 \u2007", "1 \u2007 \u2007", "1 2/7", "2 \u2007 \u2007", "0 \u2007 \u2007",
           "0 \u2007\u2007 \u2007\u2007", "1 \u2007\u2007 \u2007\u2007", "1 \u20073/10", "1 98/99", " \u20071/99",
           "0 \u2007 ", "1 \u2007 ", "1 2/8", "2 \u2007 ", "0 \u2007 "]
COMMON += ["0/1", "1/1", "9/7", "2/1", "0/1"] * 2
CASES += [(123.456, "# ?/16"), (123.456, "# ?/10"), (123.456, "# ??/100")]
COMMON += ["123 7/16", "123 5/10", "123 46/100"]


def module(name, path):
    spec = importlib.util.spec_from_file_location(name, ROOT / path)
    result = importlib.util.module_from_spec(spec); spec.loader.exec_module(result); return result


fixture = module("independent_format_inputs", "scripts/generate-xlsx-fixtures.py")
shared = module("office_format_probe", "scripts/pptx-reference.py")


def generate(date1904):
    book = fixture.Book(date1904=date1904); sheet = book.sheet("Formats"); namespace = "{" + fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"])
    formats = styles.find(namespace + "numFmts"); formats.clear(); formats.set("count", str(len(CASES)))
    cells = styles.find(namespace + "cellXfs"); cells.clear(); cells.set("count", str(len(CASES) + 1))
    base = {"numFmtId": "0", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0"}
    ET.SubElement(cells, namespace + "xf", base)
    for index, (value, code) in enumerate(CASES):
        ET.SubElement(formats, namespace + "numFmt", {"numFmtId": str(164 + index), "formatCode": code})
        ET.SubElement(cells, namespace + "xf", {**base, "numFmtId": str(164 + index), "applyNumberFormat": "1"})
        sheet.add("A" + str(index + 1), value, style=index + 1)
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    return fixture.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive: root = ET.fromstring(archive.read("content.xml"))
    def content(node):
        if node.tag == f"{{{TEXT}}}s": return " " * int(node.get(f"{{{TEXT}}}c", "1"))
        if node.tag == f"{{{TEXT}}}tab": return "\t"
        if node.tag == f"{{{TEXT}}}line-break": return "\n"
        return (node.text or "") + "".join(content(child) + (child.tail or "") for child in node)
    values = []
    for cell in root.iter(f"{{{T}}}table-cell"):
        if not cell.get(f"{{{O}}}value-type"): continue
        values.append(content(cell))
    return values


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert shared.run(["libreoffice", "--version"]) == BUILD
    cache = ROOT / ".cache/xlsx-number-formats"; cache.mkdir(parents=True, exist_ok=True); cases = []
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for date1904 in [False, True]:
            system = "1904" if date1904 else "1900"; source = work / (system + ".xlsx"); source.write_bytes(generate(date1904))
            shared.run(["libreoffice", profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
            observed = observe(source.with_suffix(".ods"))
            dates = ["1904-01-01", "1904-01-02", "1904-02-29", "1904-03-01", "1904-03-02"] if date1904 else ["1899-12-30", "1899-12-31", "1900-02-27", "1900-02-28", "1900-03-01"]
            assert observed == dates + COMMON, observed
            cases.append({"dateSystem": system, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "samples": [{"value": value, "code": code, "text": text} for (value, code), text in zip(CASES, observed)]})
    report = {"scope": "Native Calc formatted text only. Historical 1900 dates differ from Excel; SDK date semantics await user decision. Number-format rendering, fitting, fonts and Excel/WPS remain unverified.", "office": BUILD, "archiveSha256": ARCHIVE_HASH, "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native number-format observations changed"
    print(f"Verified {len(CASES) * 2} formatted cells across two date systems with pinned Calc")


if __name__ == "__main__": main()
