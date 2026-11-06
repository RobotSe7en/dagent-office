"""Independent accounting underline types and Calc row/format observations.

SDK display uses the user's full-cell policy; Calc's character-width underline
is separate evidence and does not establish the chosen display or row geometry.
"""
from pathlib import Path
from io import BytesIO
from zipfile import ZipFile
from hashlib import sha256
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('accounting_rows', ROOT / 'scripts/check-xlsx-row-heights.py')
rows = importlib.util.module_from_spec(spec); spec.loader.exec_module(rows)
S = '{' + rows.native.fixture.S + '}'
FIXTURES = ROOT / 'packages/office/tests/fixtures'
REPORT = ROOT / 'docs/evidence/xlsx-accounting-underlines.json'


def inputs():
    base = dict(group='auto-height', rtl=False, value='Ab', size=11, horizontal='general', vertical='bottom', wrap=False)
    cases = [{'underline': underline, 'rtl': rtl, 'value': value} for underline in ['singleAccounting', 'doubleAccounting'] for rtl in [False, True] for value in ['Ab', 12]]
    cases += [dict(name='plain-text'), dict(name='plain-number', value=12)]
    for underline in ['singleAccounting', 'doubleAccounting']:
        cases += [{'underline': underline, **mode} for mode in [dict(wrap=True, value='Alpha beta gamma delta'), dict(rotation=45), dict(rotation=135), dict(rotation=255, value='ABC'), dict(merge='B2:C2')]]
    return [{**base, **case} for case in cases]


def generate(cases): return rows.generate(cases)


def observe(data):
    result = []
    with ZipFile(BytesIO(data)) as archive:
        styles = ET.fromstring(archive.read('xl/styles.xml')); formats = styles.find(S + 'cellXfs'); fonts = styles.find(S + 'fonts')
        for index in range(1, len(inputs()) + 1):
            sheet = ET.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml')); cell = sheet.find(S + 'sheetData/' + S + 'row/' + S + 'c[@r="B2"]')
            font = fonts[int(formats[int(cell.get('s', '0'))].get('fontId', '0'))]; underline = font.find(S + 'u')
            result.append(dict(underline='none' if underline is None else underline.get('val', 'single'), rows=[row.attrib for row in sheet.findall(S + 'sheetData/' + S + 'row')]))
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args(); native = rows.native
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']; os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    baseline = json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts']; font = next(item for item in baseline if item['file'] == 'LiberationSans-Regular.ttf')
    family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', 'Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and Path(filename).name == font['file'] and sha256(Path(filename).read_bytes()).hexdigest() == font['sha256']
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'accounting.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        producer = (output / source.name).read_bytes(); source_records = observe(data); producer_records = observe(producer)
        report = dict(scope=__doc__, office=native.BUILD, archiveSha256=native.ARCHIVE_HASH, fontManifest=[font], sourceSha256=sha256(data).hexdigest(),
                      cases=[dict(**case, source=source_records[i], producer=producer_records[i]) for i, case in enumerate(cases)],
                      limits=['Calc rewrites accounting underline types to ordinary single/double. Source retention and producer display are checked separately.', 'The SDK reserves a font-relative decoration band for full-cell underlines; Calc row heights are not substituted for measured SDK ink.', 'Existing independent cell-flow PDF vectors record Calc character-width underlines. Excel/WPS machine behavior remains unverified.'])
        if args.record:
            (FIXTURES / 'xlsx-accounting-underlines.xlsx').write_bytes(data); (FIXTURES / 'xlsx-accounting-underlines-producer.xlsx').write_bytes(producer); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-accounting-underlines.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(cases)} independent accounting sources and Calc format/row observations')


if __name__ == '__main__': main()
