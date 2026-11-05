"""Independent fixed-font cell-flow probes; PDF observations are not SDK or Excel/WPS acceptance."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('flow_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-cell-flow.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-cell-flow.json'


def cases():
    result = []
    for rtl in [False, True]:
        for obstruction in ['empty', 'value', 'styled-blank', 'formula-empty', 'merge']:
            result.append(dict(group='overflow', rtl=rtl, obstruction=obstruction, value='Alpha beta gamma delta epsilon', horizontal='general', vertical='bottom', wrap=False))
        result.append(dict(group='continuous', rtl=rtl, value='Ab', horizontal='centerContinuous', vertical='bottom', wrap=False))
        for horizontal, value in [('fill', 'Ab'), ('fill', 12), ('justify', 'Alpha beta gamma delta'), ('distributed', 'ABC'), ('distributed', 'Alpha beta')]:
            for wrap in [False, True]: result.append(dict(group='horizontal', rtl=rtl, value=value, horizontal=horizontal, vertical='bottom', wrap=wrap))
        for vertical in ['justify', 'distributed']:
            for value in ['Ab', 'Alpha\nbeta\ngamma']:
                result.append(dict(group='vertical', rtl=rtl, value=value, horizontal='left', vertical=vertical, wrap=True))
        for underline in ['singleAccounting', 'doubleAccounting']:
            for value in ['Ab', 12]: result.append(dict(group='underline', rtl=rtl, value=value, horizontal='general', vertical='bottom', wrap=False, underline=underline))
    for value, size, wrap in [('Alpha beta gamma delta epsilon', 11, True), ('Alpha\nbeta\ngamma', 11, False), ('Ab', 20, False)]:
        result.append(dict(group='auto-height', rtl=False, value=value, size=size, horizontal='general', vertical='bottom', wrap=wrap))
    return result


def generate(inputs):
    book = native.fixture.Book(); ns = '{' + native.fixture.S + '}'
    styles = ET.fromstring(book.parts['xl/styles.xml']); fonts = styles.find(ns + 'fonts'); borders = styles.find(ns + 'borders'); xfs = styles.find(ns + 'cellXfs')
    fonts.clear(); borders.clear(); xfs.clear()
    font = ET.SubElement(fonts, ns + 'font'); ET.SubElement(font, ns + 'sz', {'val': '11'}); ET.SubElement(font, ns + 'name', {'val': 'Liberation Sans'}); ET.SubElement(font, ns + 'color', {'rgb': 'FF000000'})
    ET.SubElement(borders, ns + 'border'); border = ET.SubElement(borders, ns + 'border')
    for side in ['left', 'right', 'top', 'bottom']: ET.SubElement(ET.SubElement(border, ns + side, {'style': 'thin'}), ns + 'color', {'rgb': 'FFFF0000'})
    base = {'fontId': '0', 'fillId': '0', 'numFmtId': '0', 'borderId': '0', 'xfId': '0'}
    ET.SubElement(xfs, ns + 'xf', base); ET.SubElement(xfs, ns + 'xf', {**base, 'borderId': '1', 'applyBorder': '1'})
    for index, case in enumerate(inputs, 1):
        sheet = book.sheet('Flow' + str(index)); sheet.print_area = '$A$1:$F$4'
        sheet.before = '<sheetViews><sheetView workbookViewId="0" rightToLeft="' + str(int(case['rtl'])) + '"/></sheetViews><sheetFormatPr defaultRowHeight="15" defaultColWidth="8.43"/>'
        sheet.add('A1', 'Origin', style=1); sheet.add('A3', 'After', style=1)
        if case['group'] != 'auto-height': sheet.row_props[2] = {'ht': 60, 'customHeight': 1}
        font = ET.SubElement(fonts, ns + 'font'); ET.SubElement(font, ns + 'sz', {'val': str(case.get('size', 11))}); ET.SubElement(font, ns + 'name', {'val': 'Liberation Sans'}); ET.SubElement(font, ns + 'color', {'rgb': 'FF000000'})
        if case.get('underline'): ET.SubElement(font, ns + 'u', {'val': case['underline']})
        node = ET.SubElement(xfs, ns + 'xf', {**base, 'fontId': str(len(fonts) - 1), 'applyFont': '1', 'applyAlignment': '1'})
        ET.SubElement(node, ns + 'alignment', {'horizontal': case['horizontal'], 'vertical': case['vertical'], 'wrapText': str(int(case['wrap']))})
        style = len(xfs) - 1; sheet.add('B2', case['value'], style=style)
        if case['group'] == 'continuous':
            for column in ['C', 'D']: sheet.add(column + '2', None, style=style)
        obstruction = case.get('obstruction')
        if obstruction == 'value': sheet.add('C2', 'X')
        if obstruction == 'styled-blank': sheet.add('C2', None, style=1)
        if obstruction == 'formula-empty': sheet.add('C2', '', kind='str', formula='""')
        if obstruction == 'merge': sheet.after += '<mergeCells count="1"><mergeCell ref="C2:D2"/></mergeCells>'; sheet.add('C2', None)
    for node in [fonts, borders, xfs]: node.set('count', str(len(node)))
    book.parts['xl/styles.xml'] = ET.tostring(styles, encoding='unicode')
    return native.fixture.zipped(book.finish())


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    os.environ.update({'LC_ALL': 'C.UTF-8', 'LANG': 'C.UTF-8', 'TZ': 'UTC'})
    environment = native.shared.environment(); assert environment['libreOffice'] == native.BUILD
    fonts = environment.pop('fonts'); environment['fontCount'] = len(fonts); environment['fontManifestSha256'] = sha256(json.dumps(fonts, sort_keys=True).encode()).hexdigest()
    inputs = cases(); data = generate(inputs)
    if args.record: FIXTURE.write_bytes(data)
    else: assert FIXTURE.read_bytes() == data
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'flow.xlsx'; source.write_bytes(data)
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'pdf:calc_pdf_Export', '--outdir', str(work), str(source)])
        pages = ET.fromstring(native.shared.run(['pdftotext', '-bbox', str(source.with_suffix('.pdf')), '-'])).findall('.//{http://www.w3.org/1999/xhtml}page'); assert len(pages) == len(inputs)
        for index, (page, case) in enumerate(zip(pages, inputs), 1):
            case['words'] = [{'text': word.text, 'bounds': [float(word.get(key)) for key in ['xMin', 'yMin', 'xMax', 'yMax']]} for word in page.findall('{http://www.w3.org/1999/xhtml}word')]
            svg = work / ('page-' + str(index) + '.svg'); native.shared.run(['pdftocairo', '-f', str(index), '-l', str(index), '-svg', str(source.with_suffix('.pdf')), str(svg)])
            case['lines'] = []
            for path in ET.fromstring(svg.read_bytes()).iter('{http://www.w3.org/2000/svg}path'):
                points = re.fullmatch(r'M ([\d.-]+) ([\d.-]+) L ([\d.-]+) ([\d.-]+)\s*', path.get('d', ''))
                if not points or not path.get('stroke'): continue
                transform = re.fullmatch(r'matrix\(([^)]+)\)', path.get('transform', '')); assert transform
                a, b, c, d, e, f = map(float, transform[1].split(',')); assert b == c == 0
                x1, y1, x2, y2 = map(float, points.groups())
                case['lines'].append({'stroke': path.get('stroke'), 'width': path.get('stroke-width'), 'bounds': [round(x1*a+e, 6), round(y1*d+f, 6), round(x2*a+e, 6), round(y2*d+f, 6)]})
    report = {'scope': __doc__, 'archiveSha256': native.ARCHIVE_HASH, 'environment': environment, 'sourceSha256': sha256(data).hexdigest(), 'cases': inputs}
    expected = json.loads(REPORT.read_text()) if REPORT.exists() else None
    # This fixed producer has two observed layout outcomes for the same explicit
    # newline/no-wrap/automatic-height source. Only independently recorded exact
    # alternatives are accepted; no other case or geometry gets normalized.
    if expected and expected['sourceSha256'] == report['sourceSha256'] and expected.get('producerVariants'):
        report['producerVariants'] = expected['producerVariants']
        for variant in expected['producerVariants']:
            index = variant['case'] - 1
            if report['cases'][index] not in variant['observations']:
                (ROOT / '.cache/xlsx-cell-flow-observed.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
                raise AssertionError('An unrecorded native cell-flow variant was observed')
            report['cases'][index] = expected['cases'][index]
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    elif expected != report:
        (ROOT / '.cache/xlsx-cell-flow-observed.json').write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        raise AssertionError('Native cell-flow observations changed; inspect .cache/xlsx-cell-flow-observed.json')
    print('Verified', len(inputs), 'independent native cell-flow observations')


if __name__ == '__main__': main()
