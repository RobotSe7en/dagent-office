"""Independent native category grid positions and crossing states.

Normalized PDF vector positions do not claim whole-chart auto-layout or pixels.
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
spec = importlib.util.spec_from_file_location('grid_colors', ROOT / 'scripts/check-xlsx-chart-colors.py')
colors = importlib.util.module_from_spec(spec); spec.loader.exec_module(colors)
C, A = colors.C, colors.A
FIXTURES = colors.FIXTURES
REPORT = ROOT / 'docs/evidence/xlsx-chart-grid.json'


def inputs():
    return ([{'kind': kind, 'between': between, 'reverse': reverse} for kind in ['column', 'line'] for between in ['between', 'midCat'] for reverse in [False, True]]
            + [{'kind': kind, 'between': 'between', 'reverse': False, 'delete': True} for kind in ['column', 'line']]
            + [{'kind': kind, 'between': 'midCat', 'reverse': False, 'single': True} for kind in ['column', 'line']])


def style(parent, color):
    source = ET.SubElement(parent, C + 'spPr'); line = ET.SubElement(source, A + 'ln', {'w': '12700'})
    ET.SubElement(ET.SubElement(line, A + 'solidFill'), A + 'srgbClr', {'val': color})


def generate(cases):
    data = colors.generate([{'kind': case['kind'], 'vary': False, 'series': 1} for case in cases])
    with ZipFile(BytesIO(data)) as archive: parts = {path: archive.read(path) for path in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; root = ET.fromstring(parts[path]); chart = root.find(C + 'chart'); plot = chart.find(C + 'plotArea')
        for old in chart.findall(C + 'legend'): chart.remove(old)
        category = plot.find(C + 'catAx'); value = plot.find(C + 'valAx')
        for axis, color in [(category, 'FF0000'), (value, '0000FF')]:
            for tag in ['majorTickMark', 'minorTickMark', 'tickLblPos', 'majorGridlines', 'spPr']:
                for old in axis.findall(C + tag): axis.remove(old)
            # Reinsert in the independent official XSD order before crossAx.
            at = next(i for i, node in enumerate(axis) if node.tag == C + 'crossAx')
            for tag in ['majorTickMark', 'minorTickMark', 'tickLblPos']:
                axis.insert(at, ET.Element(C + tag, {'val': 'none'})); at += 1
            appearance = ET.Element(C + 'spPr'); line = ET.SubElement(appearance, A + 'ln', {'w': '25400'})
            ET.SubElement(ET.SubElement(line, A + 'solidFill'), A + 'srgbClr', {'val': color}); axis.insert(at, appearance)
        category.find(C + 'scaling/' + C + 'orientation').set('val', 'maxMin' if case['reverse'] else 'minMax')
        if case.get('delete'): category.find(C + 'delete').set('val', '1')
        grid = ET.Element(C + 'majorGridlines'); style(grid, '00FF00')
        category.insert(next(i for i, node in enumerate(category) if node.tag == C + 'axPos') + 1, grid)
        cross = value.find(C + 'crossBetween')
        if cross is None: cross = ET.SubElement(value, C + 'crossBetween')
        cross.set('val', case['between'])
        for old in value.findall(C + 'majorUnit'): value.remove(old)
        ET.SubElement(value, C + 'majorUnit', {'val': '1'})
        scaling = value.find(C + 'scaling')
        for old in list(scaling):
            if old.tag in [C + 'min', C + 'max']: scaling.remove(old)
        ET.SubElement(scaling, C + 'max', {'val': '4'}); ET.SubElement(scaling, C + 'min', {'val': '0'})
        style(plot, '222222')
        if case.get('single'):
            for source in plot.findall('.//' + C + 'cat/' + C + 'strRef') + plot.findall('.//' + C + 'val/' + C + 'numRef'):
                formula = source.find(C + 'f'); formula.text = formula.text.replace('$40:$A$42', '$40:$A$40').replace('$40:$B$42', '$40:$B$40')
                cache = list(source)[1]; cache.find(C + 'ptCount').set('val', '1')
                for point in cache.findall(C + 'pt'):
                    if point.get('idx') != '0': cache.remove(point)
        parts[path] = ET.tostring(root, encoding='unicode')
    return colors.native.fixture.zipped(parts)


def observe(pdf, resaved, cases, work):
    result = []
    for index, case in enumerate(cases, 1):
        svg = work / f'page-{index}.svg'; colors.native.shared.run(['pdftocairo', '-f', str(index), '-l', str(index), '-svg', str(pdf), str(svg)])
        frame, lines = None, []
        def walk(node, inherited=False):
            nonlocal frame
            if node.tag.rsplit('}', 1)[-1] == 'path' and node.get('stroke'):
                rgb = re.fullmatch(r'rgb\(([\d.]+)%, ([\d.]+)%, ([\d.]+)%\)', node.get('stroke'))
                color = ''.join(f'{round(float(value)*2.55):02x}' for value in rgb.groups()) if rgb else ''
                if color in ['222222', '00ff00']:
                    assert not inherited, 'Selected PDF path has an unhandled ancestor transform'
                    transform = re.fullmatch(r'matrix\(([^)]+)\)', node.get('transform', '')); assert transform
                    a, b, c, d, e, f = map(float, transform[1].split(',')); assert b == c == 0
                    source = node.get('d'); assert re.fullmatch(r'[MLZ\d.eE+\s-]+', source)
                    points = [(float(x)*a+e, float(y)*d+f) for x, y in re.findall(r'[ML] ([\d.eE+-]+) ([\d.eE+-]+)', source)]
                    if color == '222222':
                        assert frame is None; frame = [min(x for x,y in points), min(y for x,y in points), max(x for x,y in points), max(y for x,y in points)]
                    else:
                        assert len(points) == 2 and abs(points[0][0]-points[1][0]) < .000001; lines.append(points)
            for child in node: walk(child, inherited or node.get('transform') is not None)
        walk(ET.fromstring(svg.read_bytes())); assert frame and frame[2] > frame[0] and frame[3] > frame[1]
        positions = sorted(round((line[0][0]-frame[0])/(frame[2]-frame[0]), 4) for line in lines)
        for line in lines:
            assert abs(min(point[1] for point in line)-frame[1]) < .03 and abs(max(point[1] for point in line)-frame[3]) < .03
        with ZipFile(resaved) as archive:
            root = ET.fromstring(archive.read(f'xl/charts/chart{index}.xml')); axis = root.find('.//' + C + 'catAx'); value = root.find('.//' + C + 'valAx')
            producer = {'crossBetween': value.find(C + 'crossBetween').get('val'), 'grid': axis.find(C + 'majorGridlines') is not None, 'delete': axis.find(C + 'delete').get('val')}
        result.append({**case, 'pdfCategoryGridPositions': positions, 'producer': producer})
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == colors.native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert colors.native.shared.run(['libreoffice', '--version']) == colors.native.BUILD
    cases = inputs(); data = generate(cases); cache = ROOT / '.cache/xlsx-chart-grid'; cache.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'grid.xlsx'; source.write_bytes(data)
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: colors.native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        report = {'scope': __doc__, 'office': colors.native.BUILD, 'archiveSha256': colors.native.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(),
                  'cases': observe(output / 'grid.pdf', output / 'grid.xlsx', cases, work),
                  'limits': ['Calc ignores midCat for columns in these files while retaining the source flag. Normative source projection and producer display are checked separately.', 'No full-chart layout, font geometry, pixels or Excel/WPS acceptance is claimed.']}
        if args.record:
            (FIXTURES / 'xlsx-chart-grid.xlsx').write_bytes(data); (FIXTURES / 'xlsx-chart-grid-producer.xlsx').write_bytes((output / 'grid.xlsx').read_bytes())
            REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
        else:
            assert (FIXTURES / 'xlsx-chart-grid.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report, 'Pinned category-grid vector evidence changed'
    print(f'Verified {len(cases)} independent category grid and crossing cases')


if __name__ == '__main__': main()
