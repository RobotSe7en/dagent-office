"""Independent native Office-profile axis crossing and reverse-axis observations.

Normalized PDF vectors cover crossing coordinates, not whole-chart auto-layout.
The user selected Office's perpendicular-axis crossesAt and ignored axPos.
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
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('axis_grid', ROOT / 'scripts/check-xlsx-chart-grid.py')
grid = importlib.util.module_from_spec(spec); spec.loader.exec_module(grid)
C = grid.C
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
FIXTURES = grid.FIXTURES
REPORT = ROOT / 'docs/evidence/xlsx-chart-axes.json'


def inputs():
    common = [{'categoryCrosses': c, 'valueCrosses': v} for c in ['autoZero', 'min', 'max'] for v in ['autoZero', 'min', 'max']]
    common += [{'categoryAt': -2.5}, {'categoryAt': 2, 'reverseValues': True}, {'categoryAt': 6}, {'categoryAt': -6},
               {'valueAt': 2}, {'valueAt': 2, 'between': 'midCat'}, {'valueAt': 2, 'reverse': True},
               {'categoryAt': 1.5, 'valueAt': 2, 'reverse': True, 'reverseValues': True},
               {'categoryAt': 0, 'valueAt': 2, 'hiddenFirst': True},
               {'categoryAt': 0, 'categoryPosition': 't', 'valuePosition': 'r'},
               {'categoryAt': 0, 'categoryPosition': 'l', 'valuePosition': 'b'},
               {'minimum': 1, 'maximum': 4}, {'minimum': -4, 'maximum': -1}]
    return [{'kind': kind, 'between': 'between', 'reverse': False, **case} for kind in ['column', 'line'] for case in common]


def generate(cases):
    with ZipFile(BytesIO(grid.generate(cases))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; root = ET.fromstring(parts[path]); category = root.find('.//' + C + 'catAx'); value = root.find('.//' + C + 'valAx')
        scaling = value.find(C + 'scaling'); scaling.find(C + 'min').set('val', str(case.get('minimum', -4))); scaling.find(C + 'max').set('val', str(case.get('maximum', 4)))
        scaling.find(C + 'orientation').set('val', 'maxMin' if case.get('reverseValues') else 'minMax')
        for axis, key in [(category, 'category'), (value, 'value')]:
            for old in axis.findall(C + 'crosses') + axis.findall(C + 'crossesAt'): axis.remove(old)
            at = next(i for i, node in enumerate(axis) if node.tag == C + 'crossAx') + 1
            tag = 'crossesAt' if key + 'At' in case else 'crosses'
            setting = case.get(key + 'At', case.get(key + 'Crosses', 'autoZero'))
            axis.insert(at, ET.Element(C + tag, {'val': str(setting)}))
            if key + 'Position' in case: axis.find(C + 'axPos').set('val', case[key + 'Position'])
        if case.get('hiddenFirst'):
            sheet_path = f'xl/worksheets/sheet{index}.xml'; sheet = ET.fromstring(parts[sheet_path]); sheet.find('.//' + S + 'row[@r="40"]').set('hidden', '1'); parts[sheet_path] = ET.tostring(sheet, encoding='unicode')
            root.find(C + 'chart/' + C + 'plotVisOnly').set('val', '1')
        parts[path] = ET.tostring(root, encoding='unicode')
    return grid.colors.native.fixture.zipped(parts)


def native_axes(archive, index):
    root = ET.fromstring(archive.read(f'xl/charts/chart{index}.xml'))
    def properties(name):
        axis = root.find('.//' + C + name)
        return {node.tag[len(C):]: node.get('val') for node in axis if node.tag in [C + 'crosses', C + 'crossesAt', C + 'axPos']}
    return {'category': properties('catAx'), 'value': properties('valAx')}


def observe(pdf, producer, cases, work):
    result = []
    with ZipFile(producer) as archive:
        for index, case in enumerate(cases, 1):
            svg = work / f'page-{index}.svg'
            grid.colors.native.shared.run(['pdftocairo', '-f', str(index), '-l', str(index), '-svg', str(pdf), str(svg)])
            frame, paths = None, []
            def walk(node, inherited=False):
                nonlocal frame
                if node.tag.rsplit('}', 1)[-1] == 'path' and node.get('stroke'):
                    rgb = re.fullmatch(r'rgb\(([\d.]+)%, ([\d.]+)%, ([\d.]+)%\)', node.get('stroke'))
                    color = ''.join(f'{round(float(v)*2.55):02x}' for v in rgb.groups()) if rgb else ''
                    if color in ['222222', 'ff0000', '0000ff']:
                        assert not inherited, 'Unhandled PDF ancestor transform'
                        transform = re.fullmatch(r'matrix\(([^)]+)\)', node.get('transform', '')); assert transform
                        a,b,c,d,e,f = map(float, transform[1].split(',')); assert b == c == 0
                        path = node.get('d'); assert re.fullmatch(r'[MLZ\d.eE+\s-]+', path)
                        points = [(float(x)*a+e, float(y)*d+f) for x,y in re.findall(r'[ML] ([\d.eE+-]+) ([\d.eE+-]+)', path)]
                        if color == '222222':
                            assert frame is None
                            frame = [min(x for x,y in points), min(y for x,y in points), max(x for x,y in points), max(y for x,y in points)]
                        else: paths.append((color, points))
                for child in node: walk(child, inherited or node.get('transform') is not None)
            walk(ET.fromstring(svg.read_bytes())); assert frame and len(paths) == 2
            axes = {}
            for color, points in paths:
                assert len(points) == 2
                normalized = [[round((x-frame[0])/(frame[2]-frame[0]), 4), round((y-frame[1])/(frame[3]-frame[1]), 4)] for x,y in points]
                axes['category' if color == 'ff0000' else 'value'] = normalized
            result.append({**case, 'pdfAxes': axes, 'producer': native_axes(archive, index)})
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    native = grid.colors.native
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    cases = inputs(); data = generate(cases); cache = ROOT / '.cache/xlsx-chart-axes'; cache.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'axes.xlsx'; source.write_bytes(data)
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        report = {'scope': __doc__, 'office': native.BUILD, 'archiveSha256': native.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(),
                  'policy': {'axPos': 'preserved, ignored for display', 'crossesAt': 'perpendicular axis; integral category numbers, decimal values', 'categoryNumbers': 'current plotted categories, starting at 1'},
                  'cases': observe(output / 'axes.pdf', output / 'axes.xlsx', cases, work),
                  'limits': ['Calc retains midCat on columns but draws between categories; source and producer geometry are evaluated separately.',
                             'A nonintegral crossing of a category axis is outside the selected Office profile; no rounding is inferred from Calc.',
                             'Normalized axis vectors do not establish full labels, fonts, auto-layout, pixels or Excel/WPS acceptance.']}
        if args.record:
            (FIXTURES / 'xlsx-chart-axes.xlsx').write_bytes(data); (FIXTURES / 'xlsx-chart-axes-producer.xlsx').write_bytes((output / 'axes.xlsx').read_bytes())
            REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-chart-axes.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report, 'Pinned native axis evidence changed'
    print(f'Verified {len(cases)} independent native Office-profile axis cases')


if __name__ == '__main__': main()
