"""Independent chart formulas with wrong caches and pinned Calc current values."""
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
spec = importlib.util.spec_from_file_location('chart_formula_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-chart-formulas.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-chart-formulas.json'
FORMULAS = ['ChartConstant+7', 'SUM(B2,6)', 'IF(B3>10,B3*2,0)', 'AVERAGE(B2:B4)', 'ROUND(B5,2)',
            'COUNT(B2:B6)', 'LEN(A2)', 'TRUE()', '"text"', '""', '1/0', 'Second!A1-10']
EXPECTED = [14, 20, 40, 74 / 3, 24.67, 5, 6, True, 'text', '', '#DIV/0!', -7]


def generate(cached=True):
    book = native.fixture.Book(); sheet = book.sheet('Data'); book.sheet('Second').add('A1', 3)
    book.names = [({'name': 'ChartConstant'}, '7')]
    sheet.add('A1', 'Category')
    sheet.add('B1', 'CachedSeries' if cached else None, kind='str', formula='"Current"&" values"', formula_attrs={'ca': '1'})
    for index, formula in enumerate(FORMULAS, 2):
        sheet.add('A' + str(index), 'Cache' if cached else None, kind='str', formula='"Point"&' + str(index - 2), formula_attrs={'ca': '1'})
        sheet.add('B' + str(index), 777 if cached else None, formula=formula, formula_attrs={'ca': '1'})
    drawing = native.fixture.add_drawing(book, sheet, native.fixture.one_anchor(native.fixture.graphic_frame(1, 'chart'), 3, 2, 360, 240))
    chart = native.fixture.drawing.chart_xml('column', ['Cache'] * 12, [('CachedSeries', [999] * 12)])
    chart = re.sub(r'<c:externalData\b.*?</c:externalData>', '', chart)
    path = 'xl/charts/chart1.xml'; book.part(path, chart, 'chart'); book.relate(drawing, 'chart', native.fixture.R + '/chart', path)
    parts = book.finish(); parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return native.fixture.zipped(parts)


def observe(path):
    with ZipFile(path) as archive:
        root = ET.fromstring(archive.read('xl/worksheets/sheet1.xml'))
        shared = [''.join(item.itertext()) for item in ET.fromstring(archive.read('xl/sharedStrings.xml'))] if 'xl/sharedStrings.xml' in archive.namelist() else []
    result = {}
    for cell in root.findall('{*}sheetData/{*}row/{*}c'):
        address = cell.get('r'); kind = cell.get('t', 'n'); value = cell.findtext('{*}v')
        if address == 'A1': continue
        if kind == 'b': typed = {'type': 'boolean', 'value': value == '1'}
        elif kind == 'e': typed = {'type': 'error', 'value': value}
        elif kind in ['s', 'str', 'inlineStr']:
            typed = {'type': 'text', 'value': shared[int(value)] if kind == 's' else cell.findtext('{*}is/{*}t', '') if kind == 'inlineStr' else value or ''}
        else: typed = {'type': 'number', 'value': float(value)}
        result[address] = typed
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    data = generate(); reference = generate(cached=False)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'formulas.xlsx'; source.write_bytes(reference)
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        values = observe(output / 'formulas.xlsx')
    expected = {'B1': {'type': 'text', 'value': 'Current values'}}
    for index, value in enumerate(EXPECTED, 2):
        expected['A' + str(index)] = {'type': 'text', 'value': 'Point' + str(index - 2)}
        expected['B' + str(index)] = {'type': 'boolean' if isinstance(value, bool) else 'number' if isinstance(value, (int, float)) else 'error' if value == '#DIV/0!' else 'text', 'value': value}
    for address, value in expected.items():
        actual = values[address]
        assert actual['type'] == value['type'] and (abs(actual['value'] - value['value']) < 1e-12 if value['type'] == 'number' else actual['value'] == value['value']), (address, actual, value)
    report = {'scope': __doc__ + ' Data acceptance only; no Excel/WPS or visual acceptance claim.', 'office': native.BUILD,
              'archiveSha256': native.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(), 'referenceInputSha256': sha256(reference).hexdigest(),
              'referenceInput': 'Same formulas with formula caches omitted. Pinned Calc may retain wrong supplied caches despite recalc flags; those caches are not expected results.',
              'formulas': { 'B' + str(index): '=' + formula for index, formula in enumerate(FORMULAS, 2)}, 'expected': expected, 'producer': values}
    if args.record: FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else:
        assert FIXTURE.read_bytes() == data
        assert json.loads(REPORT.read_text()) == report
    print('Verified 25 independent current chart formula values against pinned Calc')


if __name__ == '__main__': main()
