"""Independent native explicit auto markers and producer observations.

The official Transitional schema permits auto; MS-OI29500 2.1.1588 leaves its
shape to the application. The user selected Office semantics for the separate
legacy lineChart/marker Boolean: retain its value, ignore it for display.
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
spec = importlib.util.spec_from_file_location('auto_markers', ROOT / 'scripts/check-xlsx-chart-markers.py')
markers = importlib.util.module_from_spec(spec); spec.loader.exec_module(markers)
colors, C = markers.colors, markers.C
FIXTURES = markers.FIXTURES
REPORT = ROOT / 'docs/evidence/xlsx-chart-auto-markers.json'


def inputs():
    return [{'series': 1, 'symbol': 'auto'},
            {'series': 1, 'symbol': 'circle', 'pointSymbol': 'auto', 'pointSize': 20},
            {'series': 1, 'symbol': 'auto', 'pointSymbol': 'diamond', 'pointSize': 20},
            {'series': 1, 'symbol': 'auto', 'pointSize': 24},
            {'series': 1, 'symbol': 'auto', 'pointPaint': 'none'},
            {'series': 2, 'symbol': 'auto'},
            {'series': 2, 'symbol': 'auto', 'pointSymbol': 'auto', 'pointSize': 20},
            {'series': 1, 'symbol': 'auto', 'pointSymbol': 'none'}]


def generate(cases):
    data = colors.generate([{'kind': 'line', 'vary': False, 'series': case['series']} for case in cases])
    with ZipFile(BytesIO(data)) as archive: parts = {path: archive.read(path) for path in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f'xl/charts/chart{index}.xml'; root = ET.fromstring(parts[path])
        for series in root.findall('.//' + C + 'ser'):
            marker = series.find(C + 'marker'); marker.find(C + 'symbol').set('val', case['symbol']); marker.find(C + 'size').set('val', '10')
            markers.paint(marker, series=True)
            if any(key in case for key in ['pointSymbol', 'pointSize', 'pointPaint']):
                point = ET.Element(C + 'dPt'); ET.SubElement(point, C + 'idx', {'val': '1'}); marker = ET.SubElement(point, C + 'marker')
                if 'pointSymbol' in case: ET.SubElement(marker, C + 'symbol', {'val': case['pointSymbol']})
                if 'pointSize' in case: ET.SubElement(marker, C + 'size', {'val': str(case['pointSize'])})
                if 'pointPaint' in case: markers.paint(marker, case['pointPaint'])
                series.insert(next(i for i, node in enumerate(series) if node.tag == C + 'cat'), point)
        parts[path] = ET.tostring(root, encoding='unicode')
    return colors.native.fixture.zipped(parts)


def legacy_flags(path):
    with ZipFile(path) as archive:
        paths = sorted([name for name in archive.namelist() if name.startswith('xl/charts/chart') and name.endswith('.xml')], key=lambda name: int(name.rsplit('chart', 1)[1][:-4]))
        return [[dict(node.attrib) for node in ET.fromstring(archive.read(name)).findall('.//' + C + 'lineChart/' + C + 'marker')] for name in paths]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == colors.native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert colors.native.shared.run(['libreoffice', '--version']) == colors.native.BUILD
    cases = inputs(); data = generate(cases); cache = ROOT / '.cache/xlsx-chart-auto-markers'; cache.mkdir(exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / 'out'; output.mkdir(); source = work / 'auto-markers.xlsx'; source.write_bytes(data)
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: colors.native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        report = {'scope': __doc__, 'office': colors.native.BUILD, 'archiveSha256': colors.native.ARCHIVE_HASH,
                  'sourceSha256': sha256(data).hexdigest(), 'cases': cases, 'sourceMarkers': markers.observe(source), 'producerMarkers': markers.observe(output / source.name),
                  'sourceLegacyFlags': legacy_flags(source), 'producerLegacyFlags': legacy_flags(output / source.name),
                  'pdfObservations': colors.observe(output / 'auto-markers.pdf', output / source.name, [{'kind': 'line', **case} for case in cases]),
                  'limits': ['Auto shape selection is application-owned, not a claim of matching Excel, WPS or Calc.',
                             'SDK source retains auto rather than replacing it with the preview shape.',
                             'Calc drops explicit auto series and point markers, including size and appearance in these probes; this is recorded as loss, not source fidelity.',
                             'Calc writes lineChart/marker=true after removing explicit auto. The SDK retains it without restoring discarded marker settings; public source and producer package reopen are checked separately.',
                             'The older lineChart/marker Boolean follows the selected Office semantics: it neither synthesizes nor suppresses explicit series or point markers.',
                             'No whole-chart layout, font geometry or pixel equivalence is claimed.']}
        if args.record:
            (FIXTURES / 'xlsx-chart-auto-markers.xlsx').write_bytes(data)
            (FIXTURES / 'xlsx-chart-auto-markers-producer.xlsx').write_bytes((output / source.name).read_bytes())
            REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-chart-auto-markers.xlsx').read_bytes() == data
            assert markers.observe(FIXTURES / 'xlsx-chart-auto-markers-producer.xlsx') == report['producerMarkers']
            assert json.loads(REPORT.read_text()) == report, 'Pinned automatic marker observations changed'
    print(f'Verified {len(cases)} independent explicit auto marker cases')


if __name__ == '__main__': main()
