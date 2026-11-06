"""Record current-source chart number-format behavior with pinned Calc.

The independent XLSX has data outside the print area and deliberately different
chart cache formats. Observations are producer evidence, not Excel/WPS claims.
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
spec = importlib.util.spec_from_file_location("chart_format_native", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-chart-format-semantics.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-chart-formats.xlsx"
C = "{" + native.fixture.C + "}"
S = "{" + native.fixture.S + "}"
HTML = "{http://www.w3.org/1999/xhtml}"


def inputs():
    return ([{"role": role, "linked": linked} for role in ["value-axis", "category-axis", "label", "label-percent"] for linked in [False, True, "omitted", "no-format"]]
            + [{"role": "value-axis", "linked": "no-format", "variant": variant} for variant in ["first-general", "first-blank-percent", "first-hidden-percent"]]
            + [{"role": "label", "linked": False, "variant": "point-omitted"}])


def generate(cases):
    book = native.fixture.Book()
    styles = ET.fromstring(book.parts["xl/styles.xml"])
    formats = styles.find(S + "numFmts"); cells = styles.find(S + "cellXfs")
    style_id = len(cells)
    ET.SubElement(formats, S + "numFmt", {"numFmtId": "200", "formatCode": "0.0%"}); formats.set("count", str(len(formats)))
    ET.SubElement(cells, S + "xf", {"numFmtId": "200", "fontId": "0", "fillId": "0", "borderId": "0", "xfId": "0", "applyNumberFormat": "1"}); cells.set("count", str(len(cells)))
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    for index, case in enumerate(cases, 1):
        sheet = book.sheet("Case" + str(index)); sheet.print_area = "$A$1:$L$28"
        sheet.add("B39", "Values")
        for row, category, value in [(40, .1, .25), (41, .2, .5), (42, .3, .75)]:
            sheet.add("A" + str(row), category, style=style_id)
            variant = case.get("variant")
            sheet.add("B" + str(row), None if row == 40 and variant == "first-blank-percent" else value, style=0 if row == 40 and variant == "first-general" or row > 40 and variant in ["first-blank-percent", "first-hidden-percent"] else style_id)
        if case.get("variant") == "first-hidden-percent": sheet.row_props = {40: {"hidden": 1}}
        frame = native.fixture.graphic_frame(1, "chart")
        drawing = native.fixture.add_drawing(book, sheet, native.fixture.one_anchor(frame, 0, 0, 480, 300))
        source = native.fixture.drawing.chart_xml("pie" if case["role"] == "label-percent" else "column", [.1, .2, .3], [("Values", [.25, .5, .75])])
        source = re.sub(r'<c:externalData\b.*?</c:externalData>', '', source).replace("Data!", sheet.name + "!")
        source = source.replace("$B$1", "$B$39").replace("$A$2:$A$4", "$A$40:$A$42").replace("$B$2:$B$4", "$B$40:$B$42")
        # Numeric category data allows category-axis format overrides to be
        # distinguished from an ordinary string category.
        category = re.search(r'<c:cat>.*?</c:cat>', source).group(0).replace("strRef", "numRef").replace("strCache", "numCache").replace('<c:numCache>', '<c:numCache><c:formatCode>0.00000</c:formatCode>')
        source = re.sub(r'<c:cat>.*?</c:cat>', lambda _: category, source).replace('<c:formatCode>General</c:formatCode>', '<c:formatCode>0.00000</c:formatCode>')
        source = source.replace('<c:valAx><c:axId val="200"/><c:scaling><c:orientation val="minMax"/></c:scaling>', '<c:valAx><c:axId val="200"/><c:scaling><c:orientation val="minMax"/><c:max val="1"/><c:min val="0"/></c:scaling>')
        source = source.replace('</c:valAx>', '<c:majorUnit val="0.25"/></c:valAx>')
        linked = case["linked"]
        fmt = '' if linked == "no-format" else '<c:numFmt formatCode="0.00"' + ('' if linked == "omitted" else f' sourceLinked="{int(linked)}"') + '/>'
        if case["role"] == "value-axis": source = source.replace('<c:numFmt formatCode="0" sourceLinked="0"/>', fmt)
        elif case["role"] == "category-axis": source = source.replace('<c:axPos val="b"/>', '<c:axPos val="b"/>' + fmt)
        else:
            labels = '<c:dLbls>' + fmt + '<c:dLblPos val="outEnd"/><c:showLegendKey val="0"/><c:showVal val="' + ('0' if case['role'] == 'label-percent' else '1') + '"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="' + ('1' if case['role'] == 'label-percent' else '0') + '"/></c:dLbls>'
            if case.get('variant') == 'point-omitted': labels = labels.replace('<c:dLbls>', '<c:dLbls><c:dLbl><c:idx val="1"/><c:numFmt formatCode="0.000"/></c:dLbl>')
            source = source.replace('</c:pieChart>', labels + '</c:pieChart>') if case['role'] == 'label-percent' else source.replace('<c:axId val="100"/><c:axId val="200"/>', labels + '<c:axId val="100"/><c:axId val="200"/>')
        path = f"xl/charts/chart{index}.xml"; book.part(path, source, "chart"); book.relate(drawing, "chart", native.fixture.R + "/chart", path)
    return native.fixture.zipped(book.finish())


def observe(pdf, resaved, cases):
    pages = ET.fromstring(native.shared.run(["pdftotext", "-bbox", str(pdf), "-"])).findall(".//" + HTML + "page")
    assert len(pages) == len(cases), len(pages)
    observations = []
    with ZipFile(resaved) as archive:
        for index, (case, page) in enumerate(zip(cases, pages), 1):
            chart = ET.fromstring(archive.read(f"xl/charts/chart{index}.xml"))
            labels = [node.attrib for node in chart.findall('.//' + C + 'dLbls/' + C + 'numFmt')]
            axes = {role: [node.attrib for node in chart.findall('.//' + C + tag + '/' + C + 'numFmt')] for role, tag in [('category', 'catAx'), ('value', 'valAx')]}
            words = [{"text": word.text, "boundsPt": [float(word.get(key)) for key in ['xMin', 'yMin', 'xMax', 'yMax']]} for word in page.findall(HTML + 'word')]
            observations.append({**case, "observed": {"axisFormats": axes, "labelFormats": labels, "pointFormats": [{"index": int(node.find(C + "idx").get("val")), "format": node.find(C + "numFmt").attrib} for node in chart.findall(".//" + C + "dLbl") if node.find(C + "numFmt") is not None], "pdfWords": words}})
    return observations


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    environment = native.shared.environment(); assert environment['libreOffice'] == native.BUILD
    fonts = json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts']; assert environment['fonts'] == fonts
    cache = ROOT / '.cache/xlsx-chart-formats'; cache.mkdir(exist_ok=True); cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'chart-formats.xlsx'; source.write_bytes(data)
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        observations = observe(output / 'chart-formats.pdf', output / 'chart-formats.xlsx', cases)
    report = {"scope": "Independent native source-linked chart number-format observations. Calc producer evidence only; no Excel/WPS or SDK rendering acceptance.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(',', ':')).encode()).hexdigest()}, "cases": observations}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n'); FIXTURE.write_bytes(data)
    else:
        assert json.loads(REPORT.read_text()) == report, 'Native chart-format evidence changed'
        assert FIXTURE.read_bytes() == data, 'Independent chart-format source changed'
    print(f'Verified {len(cases)} independent chart-format cases with pinned Calc')


if __name__ == '__main__': main()
