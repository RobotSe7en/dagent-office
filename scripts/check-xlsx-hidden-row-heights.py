"""Independent hidden-column AutoFit sources and pinned Calc height observations.

SDK policy includes hidden content at its stored pre-hide width. Producer row
heights are evidence, not a claim of Excel/WPS or exact browser line metrics.
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
spec = importlib.util.spec_from_file_location('hidden_rows', ROOT / 'scripts/check-xlsx-row-heights.py')
rows = importlib.util.module_from_spec(spec); spec.loader.exec_module(rows)
FIXTURES = ROOT / 'packages/office/tests/fixtures'
REPORT = ROOT / 'docs/evidence/xlsx-hidden-row-heights.json'
S = '{' + rows.native.fixture.S + '}'


def inputs():
    base = dict(group='auto-height', rtl=False, value='Alpha beta gamma delta epsilon', size=11, horizontal='general', vertical='bottom', wrap=True)
    cases = [dict(name=f'width-{width}-hidden-{hidden}', width=width, hide=hidden) for width in [8.43, 4, 16] for hidden in [False, True]]
    cases += [dict(name='hidden-default-width', hide=True), dict(name='hidden-implicit-column-font', hide=True, columnStyle=True, value=None, size=20, wrap=False),
              dict(name='hidden-blank-font', hide=True, value=None, size=20, wrap=False), dict(name='hidden-manual-row', hide=True, manual=12),
              dict(name='hidden-row', hide=True, hidden=True), dict(name='hidden-merged', hide=True, merge='B2:C2')]
    return [{**base, **case} for case in cases]


def generate(cases):
    with ZipFile(BytesIO(rows.generate(cases))) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/worksheets/sheet{index}.xml'; root = ET.fromstring(parts[path]); previous = root.find(S + 'cols')
        style = previous[0].get('style') if previous is not None else None
        if previous is not None: root.remove(previous)
        columns = ET.Element(S + 'cols'); attributes = {'min': '2', 'max': '2', 'hidden': str(int(case['hide']))}
        if 'width' in case: attributes.update(width=str(case['width']), customWidth='1')
        if style is not None: attributes['style'] = style
        ET.SubElement(columns, S + 'col', attributes); root.insert(list(root).index(root.find(S + 'sheetData')), columns)
        parts[path] = ET.tostring(root, encoding='unicode')
    return rows.native.fixture.zipped(parts)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args(); native = rows.native
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']; os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    os.environ.update(LC_ALL='C.UTF-8', LANG='C.UTF-8', TZ='UTC'); assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    baseline = json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts']
    font = next(item for item in baseline if item['file'] == 'LiberationSans-Regular.ttf')
    family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', 'Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and Path(filename).name == font['file']
    assert sha256(Path(filename).read_bytes()).hexdigest() == font['sha256'], 'Required row-height reference font differs'
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'hidden-rows.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        with ZipFile(output / source.name) as producer:
            for index, case in enumerate(cases, 1):
                root = ET.fromstring(producer.read(f'xl/worksheets/sheet{index}.xml'))
                case['producerRows'] = [node.attrib for node in root.findall(S + 'sheetData/' + S + 'row')]
                case['producerColumns'] = [node.attrib for node in root.findall(S + 'cols/' + S + 'col')]
        report = dict(scope=__doc__, archiveSha256=native.ARCHIVE_HASH, office=native.BUILD, fontManifest=[font], sourceSha256=sha256(data).hexdigest(), cases=cases,
                      limits=['Calc may select a different width when the hidden source column omits width. The SDK uses the source sheet default.', 'No exact browser/Calc text metric or Excel/WPS machine equivalence is claimed.'])
        if args.record:
            (FIXTURES / 'xlsx-hidden-row-heights.xlsx').write_bytes(data); (FIXTURES / 'xlsx-hidden-row-heights-producer.xlsx').write_bytes((output / source.name).read_bytes()); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-hidden-row-heights.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(cases)} independent hidden-column row-height cases')


if __name__ == '__main__': main()
