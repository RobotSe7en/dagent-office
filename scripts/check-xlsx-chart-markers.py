"""Independent native chart point markers and producer preservation.

ZIP/XML sources use no SDK. Marker geometry follows ECMA-376 §21.2.3.27;
Calc's independent resave and PDF colors are observations, not pixel fidelity.
"""
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
spec = importlib.util.spec_from_file_location('marker_colors', ROOT / 'scripts/check-xlsx-chart-colors.py')
colors = importlib.util.module_from_spec(spec); spec.loader.exec_module(colors)
C, A = colors.C, colors.A
FIXTURES = colors.FIXTURES
REPORT = ROOT / 'docs/evidence/xlsx-chart-markers.json'
SYMBOLS = ['circle', 'square', 'diamond', 'triangle', 'x', 'star', 'plus', 'dot', 'dash', 'none']


def inputs():
    return ([{'marker': symbol, 'size': 20} for symbol in SYMBOLS]
            + [{'marker': 'diamond'}, {'size': 24}, {'appearance': 'solid'},
               {'marker': 'circle', 'size': 20, 'appearance': 'none'},
               {'marker': 'square', 'size': 20, 'appearance': 'zero-line'},
               {'marker': 'diamond', 'size': 20, 'appearance': 'gradient'}])


def paint(marker, kind='solid', series=False):
    style = ET.SubElement(marker, C + 'spPr')
    if kind == 'none': ET.SubElement(style, A + 'noFill')
    elif kind == 'gradient':
        gradient = ET.SubElement(style, A + 'gradFill'); stops = ET.SubElement(gradient, A + 'gsLst')
        for position, color in [('0', 'FF0000'), ('100000', '0000FF')]:
            stop = ET.SubElement(stops, A + 'gs', {'pos': position}); ET.SubElement(stop, A + 'srgbClr', {'val': color})
        ET.SubElement(gradient, A + 'lin', {'ang': '2700000', 'scaled': '0'})
    else: ET.SubElement(ET.SubElement(style, A + 'solidFill'), A + 'srgbClr', {'val': '228833' if series else 'FF0000'})
    line = ET.SubElement(style, A + 'ln', {'w': '0' if kind == 'zero-line' else '19050'})
    ET.SubElement(ET.SubElement(line, A + 'solidFill'), A + 'srgbClr', {'val': '111111' if series else '0000FF'})


def generate(cases):
    source = colors.generate([{'kind': 'line', 'vary': False, 'series': 1} for _ in cases])
    with ZipFile(BytesIO(source)) as archive: parts = {path: archive.read(path) for path in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; root = ET.fromstring(parts[path]); series = root.find('.//' + C + 'ser')
        marker = series.find(C + 'marker'); marker.find(C + 'size').set('val', '10'); paint(marker, series=True)
        point = ET.Element(C + 'dPt'); ET.SubElement(point, C + 'idx', {'val': '1'}); marker = ET.SubElement(point, C + 'marker')
        if 'marker' in case: ET.SubElement(marker, C + 'symbol', {'val': case['marker']})
        if 'size' in case: ET.SubElement(marker, C + 'size', {'val': str(case['size'])})
        if 'appearance' in case: paint(marker, case['appearance'])
        at = next(i for i, node in enumerate(series) if node.tag == C + 'cat'); series.insert(at, point)
        point = ET.Element(C + 'dPt'); ET.SubElement(point, C + 'idx', {'val': '2'}); ET.SubElement(ET.SubElement(point, C + 'marker'), C + 'symbol', {'val': 'none'})
        series.insert(at + 1, point); parts[path] = ET.tostring(root, encoding='unicode')
    return colors.native.fixture.zipped(parts)


def observe(path):
    def marker(source):
        if source is None: return None
        # Attribute order and native prefixes are not semantic.
        return {'symbol': None if source.find(C + 'symbol') is None else source.find(C + 'symbol').get('val'),
                'size': None if source.find(C + 'size') is None else source.find(C + 'size').get('val'),
                'appearance': [[node.tag.rsplit('}', 1)[-1], dict(sorted(node.attrib.items()))] for node in source.iter() if node is not source and node.tag not in [C + 'symbol', C + 'size']]}
    with ZipFile(path) as archive:
        charts = sorted([name for name in archive.namelist() if name.startswith('xl/charts/chart') and name.endswith('.xml')], key=lambda name: int(name.rsplit('chart', 1)[1][:-4]))
        return [[{'series': marker(series.find(C + 'marker')),
                  'points': [{'index': point.find(C + 'idx').get('val'), 'marker': marker(point.find(C + 'marker'))} for point in series.findall(C + 'dPt')]} for series in ET.fromstring(archive.read(path)).findall('.//' + C + 'ser')] for path in charts]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == colors.native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert colors.native.shared.run(['libreoffice', '--version']) == colors.native.BUILD
    cases = inputs(); data = generate(cases); cache = ROOT / '.cache/xlsx-chart-markers'; cache.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'markers.xlsx'; source.write_bytes(data)
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: colors.native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        related = []
        for theme in ['default', 'alternate']:
            color_source = FIXTURES / f'xlsx-chart-colors-{theme}.xlsx'
            colors.native.shared.run(['libreoffice', profile, '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(color_source)])
            producer = output / color_source.name; fixture = FIXTURES / f'xlsx-chart-colors-{theme}-producer.xlsx'
            related.append({'theme': theme, 'sourceSha256': sha256(color_source.read_bytes()).hexdigest(), 'producerMarkers': observe(producer)})
            if args.record: fixture.write_bytes(producer.read_bytes())
            else: assert observe(fixture) == observe(producer)
        report = {'scope': __doc__, 'office': colors.native.BUILD, 'archiveSha256': colors.native.ARCHIVE_HASH,
                  'sourceSha256': sha256(data).hexdigest(), 'cases': cases, 'relatedColors': related, 'sourceMarkers': observe(source), 'producerMarkers': observe(output / 'markers.xlsx'),
                  'pdfObservations': colors.observe(output / 'markers.pdf', output / 'markers.xlsx', [{'kind': 'line', **case} for case in cases]),
                  'limits': ['The two related color producers contain 26 charts, including varied per-point appearance; public import and exact HTML-package reopen are checked separately.', 'No marker pixel equivalence to Calc is claimed.', 'Calc changes star to square and dot to triangle, drops the size on none, and materializes inherited properties; these changes are recorded rather than treated as source fidelity.', 'The source normative dash/dot geometry and star/plus/x fills are checked separately in all browser engines.', 'Excel/WPS unverified.']}
        if args.record:
            (FIXTURES / 'xlsx-chart-markers.xlsx').write_bytes(data)
            (FIXTURES / 'xlsx-chart-markers-producer.xlsx').write_bytes((output / 'markers.xlsx').read_bytes())
            REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-chart-markers.xlsx').read_bytes() == data
            assert observe(FIXTURES / 'xlsx-chart-markers-producer.xlsx') == report['producerMarkers']
            assert json.loads(REPORT.read_text()) == report, 'Pinned native marker observations changed'
    print(f'Verified {len(cases)} native chart marker cases and producer properties')


if __name__ == '__main__': main()
