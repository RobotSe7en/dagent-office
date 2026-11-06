"""Independent automatic/fixed row-height sources and pinned Calc observations."""
from hashlib import sha256
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('flow', ROOT / 'scripts/check-xlsx-cell-flow.py')
flow = importlib.util.module_from_spec(spec); spec.loader.exec_module(flow)
native = flow.native
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-row-heights.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-row-heights.json'
FLOOR_FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-row-font-floor.xlsx'
FLOOR_REPORT = ROOT / 'docs/evidence/xlsx-row-font-floor.json'


def cases():
    base = dict(group='auto-height', rtl=False, value='Alpha beta gamma delta epsilon', size=11, horizontal='general', vertical='bottom', wrap=True)
    return [{**base, **case} for case in [
        dict(name='wrapped'), dict(name='fixed-sheet', defaultHeight=18, customDefault=True),
        dict(name='automatic-exception', defaultHeight=18, customDefault=True, automatic=True),
        dict(name='manual-row', manual=12), dict(name='hidden-row', hidden=True),
        dict(name='offscreen-column', column='D'), dict(name='merged-columns', merge='B2:C2'),
        dict(name='merged-rows', merge='B2:B3'), dict(name='rotation-45', wrap=False, rotation=45),
        dict(name='vertical', wrap=False, rotation=255, value='ABC'), dict(name='large-font', value='Ab', size=20, wrap=False),
        dict(name='explicit-lines', value='Alpha\nbeta\ngamma', wrap=True),
        dict(name='hidden-column', hiddenColumn=True), dict(name='far-row', row=300),
        dict(name='blank-font', value=None, size=20, wrap=False),
        dict(name='row-font', value=None, size=20, wrap=False, rowStyle=True),
        dict(name='column-font', value=None, size=20, wrap=False, columnStyle=True),
        dict(name='empty-string-font', value='', size=20, wrap=False),
        dict(name='shrunk-font', value='Alpha beta gamma delta epsilon', size=20, wrap=False, shrink=True),
        dict(name='rotated-short-font', value='A', size=20, wrap=False, rotation=45),
        dict(name='hidden-display-font', value=12, size=20, wrap=False, hideDisplay=True),
    ]]


def floor_cases():
    return [dict(group='auto-height', name=f'normal20-small{size}-{kind}-default{height}', rtl=False, value=value, size=size,
                 horizontal='general', vertical='bottom', wrap=False, defaultHeight=height)
            for height in [15, 30] for size in [5, 8, 11] for kind, value in [('value', 'Ab'), ('blank', None)]]


def generate(inputs, normal_size=11):
    data = flow.generate(inputs); ns = '{' + native.fixture.S + '}'
    with ZipFile(BytesIO(data)) as source: parts = {name: source.read(name) for name in source.namelist()}
    styles = ET.fromstring(parts['xl/styles.xml']); xfs = styles.find(ns + 'cellXfs')
    styles.find(ns + 'fonts')[0].find(ns + 'sz').set('val', str(normal_size))
    for index, case in enumerate(inputs, 1):
        path = 'xl/worksheets/sheet' + str(index) + '.xml'; root = ET.fromstring(parts[path]); fmt = root.find(ns + 'sheetFormatPr')
        fmt.set('defaultRowHeight', str(case.get('defaultHeight', 15)))
        if 'customDefault' in case: fmt.set('customHeight', str(int(case['customDefault'])))
        rows = root.find(ns + 'sheetData'); row = rows.find(ns + "row[@r='2']"); cell = row.find(ns + 'c')
        if 'manual' in case: row.set('customHeight', '1'); row.set('ht', str(case['manual']))
        if case.get('automatic'): row.set('customHeight', '0'); row.set('ht', '80')
        if case.get('hidden'): row.set('hidden', '1')
        if case.get('column'): cell.set('r', case['column'] + '2')
        if case.get('rotation') is not None: xfs[int(cell.get('s'))].find(ns + 'alignment').set('textRotation', str(case['rotation']))
        if case.get('merge'):
            merge = ET.Element(ns + 'mergeCells', {'count': '1'}); ET.SubElement(merge, ns + 'mergeCell', {'ref': case['merge']}); root.insert(list(root).index(rows) + 1, merge)
        if case.get('hiddenColumn'):
            columns = ET.Element(ns + 'cols'); ET.SubElement(columns, ns + 'col', {'min': '2', 'max': '2', 'hidden': '1'}); root.insert(list(root).index(rows), columns)
        if case.get('shrink'): xfs[int(cell.get('s'))].find(ns + 'alignment').set('shrinkToFit', '1')
        if case.get('hideDisplay'):
            formats = styles.find(ns + 'numFmts'); ET.SubElement(formats, ns + 'numFmt', {'numFmtId': '9000', 'formatCode': ';;;'}); formats.set('count', str(len(formats))); xfs[int(cell.get('s'))].set('numFmtId', '9000'); xfs[int(cell.get('s'))].set('applyNumberFormat', '1')
        if case.get('rowStyle'):
            row.set('s', cell.get('s')); row.set('customFormat', '1'); row.remove(cell)
        if case.get('columnStyle'):
            columns = ET.Element(ns + 'cols'); ET.SubElement(columns, ns + 'col', {'min': '2', 'max': '2', 'style': cell.get('s')}); root.insert(list(root).index(rows), columns); row.remove(cell)
        if case.get('row'):
            row.set('r', str(case['row'])); cell.set('r', 'B' + str(case['row'])); rows.remove(row); rows.append(row)
        parts[path] = ET.tostring(root, encoding='unicode')
    parts['xl/styles.xml'] = ET.tostring(styles, encoding='unicode')
    return native.fixture.zipped(parts)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    os.environ.update({'LC_ALL': 'C.UTF-8', 'LANG': 'C.UTF-8', 'TZ': 'UTC'})
    environment = native.shared.environment(); assert environment['libreOffice'] == native.BUILD
    fonts = environment.pop('fonts'); environment['fontCount'] = len(fonts); environment['fontManifestSha256'] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest()
    for inputs, normal_size, fixture, report_path in [(cases(), 11, FIXTURE, REPORT), (floor_cases(), 20, FLOOR_FIXTURE, FLOOR_REPORT)]:
        data = generate(inputs, normal_size)
        if args.record: fixture.write_bytes(data)
        else: assert fixture.read_bytes() == data
        with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
            work = Path(temporary); source = work / 'rows.xlsx'; source.write_bytes(data); output = work / 'producer'; output.mkdir()
            native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
            with ZipFile(output / 'rows.xlsx') as producer:
                ns = '{' + native.fixture.S + '}'
                for index, case in enumerate(inputs, 1):
                    root = ET.fromstring(producer.read('xl/worksheets/sheet' + str(index) + '.xml'))
                    case['producerFormat'] = root.find(ns + 'sheetFormatPr').attrib if root.find(ns + 'sheetFormatPr') is not None else None
                    case['producerRows'] = [node.attrib for node in root.findall(ns + 'sheetData/' + ns + 'row')]
                    if normal_size == 20:
                        row = next(row for row in case['producerRows'] if row['r'] == '2')
                        assert float(row['ht']) == case['defaultHeight'], (case, row)
        report = {'scope': __doc__, 'archiveSha256': native.ARCHIVE_HASH, 'environment': environment, 'sourceSha256': sha256(data).hexdigest(), 'cases': inputs}
        if normal_size == 20: report['normalFontPt'] = normal_size
        if args.record: report_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else: assert json.loads(report_path.read_text()) == report
        print('Verified', len(inputs), 'independent native row-height sources and Calc observations')


if __name__ == '__main__': main()
