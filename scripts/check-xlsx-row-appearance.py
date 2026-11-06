"""Independent implicit rule/table row fonts and validation-only current values.

Calc observations record producer differences; they do not establish that Calc
implements custom table styles or every differential font property.
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
spec = importlib.util.spec_from_file_location('row_appearance_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-row-appearance.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-row-appearance.json'
CASES = ['conditional-match', 'conditional-miss', 'conditional-relative', 'table-custom', 'table-data-differential',
         'table-and-conditional', 'validation-current', 'color-only-whole-column', 'table-named-font']


def generate(cached=True):
    book = native.fixture.Book(); ns = '{' + native.fixture.S + '}'
    styles = ET.fromstring(book.parts['xl/styles.xml']); font = styles.find(ns + 'fonts')[0]
    font.find(ns + 'name').set('val', 'Liberation Sans'); font.find(ns + 'sz').set('val', '11')
    for child in list(font):
        if child.tag == ns + 'scheme': font.remove(child)
    fonts = styles.find(ns + 'fonts'); named_font = ET.fromstring(ET.tostring(font)); named_font.find(ns + 'sz').set('val', '20')
    font_id = len(fonts); fonts.append(named_font); fonts.set('count', str(len(fonts)))
    named_xfs = styles.find(ns + 'cellStyleXfs'); style_id = len(named_xfs)
    ET.SubElement(named_xfs, ns + 'xf', {'fontId': str(font_id), 'fillId': '0', 'numFmtId': '0', 'borderId': '0', 'applyFont': '1'})
    named_xfs.set('count', str(len(named_xfs))); named_styles = styles.find(ns + 'cellStyles')
    ET.SubElement(named_styles, ns + 'cellStyle', {'name': 'BigRowsCells', 'xfId': str(style_id)}); named_styles.set('count', str(len(named_styles)))
    dxfs = styles.find(ns + 'dxfs'); indices = []
    for size in [20, 30]:
        indices.append(len(dxfs)); ET.SubElement(ET.SubElement(ET.SubElement(dxfs, ns + 'dxf'), ns + 'font'), ns + 'sz', {'val': str(size)})
    color_index = len(dxfs); ET.SubElement(ET.SubElement(ET.SubElement(dxfs, ns + 'dxf'), ns + 'font'), ns + 'color', {'rgb': 'FFFF0000'})
    dxfs.set('count', str(len(dxfs)))
    xfs = styles.find(ns + 'cellXfs'); wrap = len(xfs)
    xf = ET.SubElement(xfs, ns + 'xf', {'fontId': '0', 'fillId': '0', 'numFmtId': '0', 'borderId': '0', 'xfId': '0', 'applyAlignment': '1'})
    ET.SubElement(xf, ns + 'alignment', {'wrapText': '1'}); xfs.set('count', str(len(xfs)))
    tables = styles.find(ns + 'tableStyles'); tables.set('count', '1')
    table_style = ET.SubElement(tables, ns + 'tableStyle', {'name': 'BigRows', 'table': '1', 'pivot': '0', 'count': '1'})
    ET.SubElement(table_style, ns + 'tableStyleElement', {'type': 'wholeTable', 'dxfId': str(indices[0])})
    book.parts['xl/styles.xml'] = ET.tostring(styles, encoding='unicode')
    for index, case in enumerate(CASES, 1):
        sheet = book.sheet('Rows' + str(index)); sheet.add('A1', 'Origin'); sheet.add('A3', 'After')
        if case.startswith('table'):
            for column in ['B', 'C', 'D']: sheet.add(column + '1', column)
            data = f' dataDxfId="{indices[0]}"' if case == 'table-data-differential' else ' dataCellStyle="BigRowsCells"' if case == 'table-named-font' else ''
            style = '' if data else '<tableStyleInfo name="BigRows" showFirstColumn="0" showLastColumn="0" showRowStripes="0" showColumnStripes="0"/>'
            table = f'<table xmlns="{native.fixture.S}" id="{index}" name="RowsTable{index}" displayName="RowsTable{index}" ref="B1:D4" totalsRowShown="0"{data}><tableColumns count="3"><tableColumn id="1" name="B"/><tableColumn id="2" name="C"/><tableColumn id="3" name="D"/></tableColumns>{style}</table>'
            path = 'xl/tables/table' + str(index) + '.xml'; book.part(path, table, 'table'); book.relate(sheet.path, 'table', native.fixture.R + '/table', path)
            sheet.tail = f'<tableParts count="1"><tablePart r:id="table"/></tableParts>'
        if case.startswith('conditional') or case == 'table-and-conditional' or case == 'color-only-whole-column':
            area = 'B2:D3' if case == 'conditional-relative' else 'B1:B1048576' if case == 'color-only-whole-column' else 'B2'
            formula = '0' if case == 'conditional-miss' else 'COLUMN()=3' if case == 'conditional-relative' else '1'
            dxf = color_index if case == 'color-only-whole-column' else indices[1] if case == 'table-and-conditional' else indices[0]
            sheet.after = f'<conditionalFormatting sqref="{area}"><cfRule type="expression" dxfId="{dxf}" priority="1"><formula>{formula}</formula></cfRule></conditionalFormatting>'
        if case == 'validation-current':
            sheet.add('B2', 'x' if cached else None, kind='str', style=wrap, formula='"Alpha beta gamma delta epsilon"', formula_attrs={'ca': '1'})
            sheet.after = '<dataValidations count="1"><dataValidation type="textLength" operator="between" sqref="B2" allowBlank="1"><formula1>1</formula1><formula2>100</formula2></dataValidation></dataValidations>'
    parts = book.finish(); parts['xl/workbook.xml'] = parts['xl/workbook.xml'].replace('fullCalcOnLoad="0"', 'fullCalcOnLoad="1" forceFullCalc="1"')
    return native.fixture.zipped(parts)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    source = generate(); reference = generate(cached=False); observations = []
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); path = work / 'appearance.xlsx'; path.write_bytes(reference)
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(path)])
        with ZipFile(output / 'appearance.xlsx') as archive:
            for index, case in enumerate(CASES, 1):
                root = ET.fromstring(archive.read('xl/worksheets/sheet' + str(index) + '.xml'))
                observations.append({'case': case, 'rows': [row.attrib for row in root.findall('{*}sheetData/{*}row')]})
    report = {'scope': __doc__, 'office': native.BUILD, 'archiveSha256': native.ARCHIVE_HASH, 'sourceSha256': sha256(source).hexdigest(),
              'referenceInputSha256': sha256(reference).hexdigest(), 'referenceInput': 'Same source with the deliberately wrong validation formula cache omitted.',
              'producer': observations, 'limits': ['Calc producer heights are observations, not expected SDK height for unsupported differential fonts or custom table styles.', 'Excel/WPS unverified.']}
    if args.record: FIXTURE.write_bytes(source); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else:
        assert FIXTURE.read_bytes() == source
        assert json.loads(REPORT.read_text()) == report
    print('Verified', len(CASES), 'independent row appearance sources and pinned Calc observations')


if __name__ == '__main__': main()
