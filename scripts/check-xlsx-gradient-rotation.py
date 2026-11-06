"""Independent rotated linear gradients, Calc normalization and PDF vectors.

Source flags/angles, producer properties and effective native paint vectors are
separate facts. They do not establish Excel/WPS behavior or SDK policy.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from io import BytesIO
from hashlib import sha256
import argparse
import importlib.util
import json
import os
import re
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('gradient_pictures', ROOT / 'scripts/check-xlsx-picture-rotation.py')
pictures = importlib.util.module_from_spec(spec); spec.loader.exec_module(pictures)
native, fixture = pictures.native, pictures.fixture
A, D = pictures.A, pictures.D
REPORT = ROOT / 'docs/evidence/xlsx-gradient-rotation.json'
FIXTURES = ROOT / 'packages/office/tests/fixtures'


def inputs():
    cases = [dict(rotation=rotation, rotates=rotates, angle=angle, scaled=scaled, flipH=False, flipV=False)
             for rotation in [0, 45, 90] for rotates in [True, False] for angle in [0, 45] for scaled in [False, True]]
    cases += [dict(rotation=45, rotates=False, angle=0, scaled=scaled, flipH=h, flipV=v)
              for h, v in [(True, False), (False, True), (True, True)] for scaled in [False, True]]
    return cases + [dict(rotation=45, rotates=None, angle=angle, scaled=False, flipH=False, flipV=False) for angle in [0, 45]]


def generate(cases):
    book = fixture.Book()
    for index, case in enumerate(cases, 1):
        sheet = book.sheet('Gradient' + str(index))
        shape = ET.fromstring('<root xmlns:xdr="' + fixture.D + '" xmlns:a="' + fixture.A + '">' + fixture.basic_shape(1, '') + '</root>')[0]
        properties = shape.find(D + 'spPr'); transform = properties.find(A + 'xfrm')
        transform.set('rot', str(case['rotation']*60000)); transform.set('flipH', str(int(case['flipH']))); transform.set('flipV', str(int(case['flipV'])))
        for old in properties.findall(A + 'solidFill'): properties.remove(old)
        gradient = ET.Element(A + 'gradFill', {} if case['rotates'] is None else {'rotWithShape': str(int(case['rotates']))}); stops = ET.SubElement(gradient, A + 'gsLst')
        for position, color in [('0', 'FF0000'), ('100000', '0000FF')]: ET.SubElement(ET.SubElement(stops, A + 'gs', {'pos': position}), A + 'srgbClr', {'val': color})
        ET.SubElement(gradient, A + 'lin', {'ang': str(case['angle']*60000), 'scaled': str(int(case['scaled']))}); properties.insert(2, gradient)
        fixture.add_drawing(book, sheet, fixture.one_anchor(ET.tostring(shape, encoding='unicode'), 2, 4, 180, 90))
    return fixture.zipped(book.finish())


def observe(data):
    with ZipFile(BytesIO(data)) as archive:
        results = []
        for index in range(1, len(inputs()) + 1):
            root = ET.fromstring(archive.read(f'xl/drawings/drawing{index}.xml')); gradient = root.find('.//' + A + 'gradFill'); transform = root.find('.//' + A + 'xfrm'); anchor = root[0]
            markers = {tag: {node.tag.split('}')[-1]: node.text for node in anchor.find(D + tag)} for tag in ['from', 'to'] if anchor.find(D + tag) is not None}
            results.append(dict(gradient=gradient.attrib, linear=gradient.find(A + 'lin').attrib, transform=transform.attrib, transformExtent=transform.find(A + 'ext').attrib,
                                anchor=anchor.tag.split('}')[-1], markers=markers, anchorExtent=anchor.find(D + 'ext').attrib if anchor.find(D + 'ext') is not None else None))
        return results


def vectors(pdf, work):
    records = []
    for page in range(1, len(inputs()) + 1):
        svg = work / (pdf.stem + '-' + str(page) + '.svg'); native.shared.run(['pdftocairo', '-f', str(page), '-l', str(page), '-svg', str(pdf), str(svg)])
        bands, frame = [], None
        def walk(node, transformed=False):
            nonlocal frame
            if node.tag.rsplit('}', 1)[-1] == 'path':
                fill = node.get('fill', ''); stroke = node.get('stroke', '')
                if fill.startswith('rgb('):
                    assert not transformed and node.get('transform') is None, 'Gradient bands have an unexpected coordinate transform'
                    point = re.match(r'M ([\d.eE+-]+) ([\d.eE+-]+)', node.get('d', '')); assert point
                    bands.append(tuple(map(float, point.groups())))
                if stroke:
                    assert not transformed and frame is None
                    rgb = re.fullmatch(r'rgb\(([\d.]+)%, ([\d.]+)%, ([\d.]+)%\)', stroke); assert rgb and [round(float(n)*2.55) for n in rgb.groups()] == [40, 116, 84]
                    matrix = re.fullmatch(r'matrix\(([^)]+)\)', node.get('transform', '')); assert matrix
                    a, b, c, d, e, f = map(float, matrix[1].split(','))
                    points = [(float(x)*a+float(y)*c+e, float(x)*b+float(y)*d+f) for x, y in re.findall(r'[ML] ([\d.eE+-]+) ([\d.eE+-]+)', node.get('d', ''))]
                    frame = [min(x for x,y in points), min(y for x,y in points), max(x for x,y in points), max(y for x,y in points)]
            for child in node: walk(child, transformed or node.get('transform') is not None)
        walk(ET.fromstring(svg.read_bytes())); assert frame and len(bands) > 3
        factor = len(bands)/(len(bands)-2)
        records.append(dict(bands=len(bands), aspect=round((frame[2]-frame[0])/(frame[3]-frame[1]), 3), normalizedVector=[round((bands[-1][axis]-bands[1][axis])*factor/(frame[axis+2]-frame[axis]), 3) for axis in [0, 1]]))
    return records


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']; os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    font = next(item for item in json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts'] if item['file'] == 'LiberationSans-Regular.ttf')
    family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', 'Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and Path(filename).name == font['file'] and sha256(Path(filename).read_bytes()).hexdigest() == font['sha256']
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'gradients.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir(); reopened = work / 'reopened'; reopened.mkdir()
        profile = '-env:UserInstallation=' + (work / 'profile').as_uri()
        for kind in ['pdf', 'xlsx:Calc MS Excel 2007 XML']: native.shared.run(['libreoffice', profile, '--headless', '--convert-to', kind, '--outdir', str(output), str(source)])
        producer = (output / source.name).read_bytes()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile-reopen').as_uri(), '--headless', '--convert-to', 'pdf', '--outdir', str(reopened), str(output / source.name)])
        source_vectors, producer_vectors = vectors(output / 'gradients.pdf', work), vectors(reopened / 'gradients.pdf', work)
        report = dict(scope=__doc__, office=native.BUILD, archiveSha256=native.ARCHIVE_HASH, fontManifest=[font], sourceSha256=sha256(data).hexdigest(),
                      specification='https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.gradientfill?view=openxml-3.0.1',
                      cases=[dict(**case, source=s, producer=p, sourcePdf=sv, producerPdf=pv) for case, s, p, sv, pv in zip(cases, observe(data), observe(producer), source_vectors, producer_vectors)],
                      limits=['Calc rewrites rotating gradients as fixed fields with adjusted angles, drops scaled and can change flip appearance on reopen. Source flags/angles remain independent.', 'Native quarter-turn geometry can change extents; transform extents and anchor markers remain recorded separately.', 'Normalized native band vectors are rounded to three decimals and do not claim whole-page pixels or Excel/WPS machine verification.'])
        if args.record:
            (FIXTURES / 'xlsx-gradient-rotation.xlsx').write_bytes(data); (FIXTURES / 'xlsx-gradient-rotation-producer.xlsx').write_bytes(producer); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2)+'\n')
        else:
            assert (FIXTURES / 'xlsx-gradient-rotation.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(cases)} independent gradient sources, Calc normalization and reopened PDF vectors')


if __name__ == '__main__': main()
