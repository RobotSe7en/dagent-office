"""Independent finite-double chart intervals and Calc numerical preservation.

Wide intervals can exceed the largest double even though both endpoints are
finite. Subnormal source values and any producer underflow remain distinct.
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
spec = importlib.util.spec_from_file_location('finite_colors', ROOT / 'scripts/check-xlsx-chart-colors.py')
colors = importlib.util.module_from_spec(spec); spec.loader.exec_module(colors)
C = colors.C; S = '{' + colors.native.fixture.S + '}'
REPORT = ROOT / 'docs/evidence/xlsx-chart-finite-scales.json'
FIXTURES = ROOT / 'packages/office/tests/fixtures'


def inputs():
    return ([dict(kind=kind, reverse=reverse, automatic=automatic, values=[-9e307, 0, 9e307]) for kind in ['line', 'column'] for reverse in [False, True] for automatic in [False, True]]
            + [dict(kind=kind, reverse=False, automatic=True, values=[0, 5e-324, 1e-323]) for kind in ['line', 'column']]
            + [dict(kind=kind, reverse=reverse, automatic=False, outside=True, values=[-9e307, .5, 9e307]) for kind in ['line', 'column'] for reverse in [False, True]])


def generate(cases):
    data = colors.generate([dict(kind=case['kind'], vary=False, series=1) for case in cases])
    with ZipFile(BytesIO(data)) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; chart = ET.fromstring(parts[path]); axis = chart.find('.//' + C + 'valAx'); scaling = axis.find(C + 'scaling')
        for old in list(scaling):
            if old.tag in [C + 'min', C + 'max']: scaling.remove(old)
        scaling.find(C + 'orientation').set('val', 'maxMin' if case['reverse'] else 'minMax')
        for tag in ['majorUnit', 'minorUnit']:
            for old in axis.findall(C + tag): axis.remove(old)
        if not case['automatic']:
            ET.SubElement(scaling, C + 'max', {'val': '1' if case.get('outside') else '9e307'}); ET.SubElement(scaling, C + 'min', {'val': '0' if case.get('outside') else '-9e307'}); ET.SubElement(axis, C + 'majorUnit', {'val': '.25' if case.get('outside') else '3e307'})
        for cache in chart.findall('.//' + C + 'val/' + C + 'numRef/' + C + 'numCache'):
            for point, value in zip(cache.findall(C + 'pt'), case['values']): point.find(C + 'v').text = str(value)
        parts[path] = ET.tostring(chart, encoding='unicode')
        path = f'xl/worksheets/sheet{index}.xml'; sheet = ET.fromstring(parts[path])
        for row, value in zip(range(40, 43), case['values']): sheet.find(S + f'sheetData/{S}row/{S}c[@r="B{row}"]/{S}v').text = str(value)
        parts[path] = ET.tostring(sheet, encoding='unicode')
    return colors.native.fixture.zipped(parts)


def observe(data):
    with ZipFile(BytesIO(data)) as archive:
        results = []
        for index in range(1, len(inputs()) + 1):
            chart = ET.fromstring(archive.read(f'xl/charts/chart{index}.xml')); axis = chart.find('.//' + C + 'valAx'); scaling = axis.find(C + 'scaling')
            sheet = ET.fromstring(archive.read(f'xl/worksheets/sheet{index}.xml'))
            results.append(dict(axis={node.tag.split('}')[-1]: node.get('val') for node in scaling}, majorUnit=axis.find(C + 'majorUnit').get('val') if axis.find(C + 'majorUnit') is not None else None,
                                values=[sheet.find(S + f'sheetData/{S}row/{S}c[@r="B{row}"]/{S}v').text for row in range(40, 43)]))
        return results


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args(); native = colors.native
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']; os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    font = next(item for item in json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts'] if item['file'] == 'LiberationSans-Regular.ttf')
    family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', 'Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and Path(filename).name == font['file'] and sha256(Path(filename).read_bytes()).hexdigest() == font['sha256']
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'scales.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        producer = (output / source.name).read_bytes()
        report = dict(scope=__doc__, office=native.BUILD, archiveSha256=native.ARCHIVE_HASH, fontManifest=[font], sourceSha256=sha256(data).hexdigest(),
                      cases=[dict(**case, source=s, producer=p) for case, s, p in zip(cases, observe(data), observe(producer))],
                      limits=['Native source and actual producer values/axes are checked independently, including numerical loss. No Excel/WPS or whole-chart visual fidelity is claimed.'])
        if args.record:
            (FIXTURES / 'xlsx-chart-finite-scales.xlsx').write_bytes(data); (FIXTURES / 'xlsx-chart-finite-scales-producer.xlsx').write_bytes(producer); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
        else:
            assert (FIXTURES / 'xlsx-chart-finite-scales.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(cases)} independent finite chart scales and Calc values')


if __name__ == '__main__': main()
