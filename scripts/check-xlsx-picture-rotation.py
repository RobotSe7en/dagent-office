"""Independent picture-fill rotation, reflection and crop sources.

Calc producer properties describe only that producer, not normative support for
rotWithShape=false. SDK geometry/pixels are checked separately in browsers.
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
import struct
import zlib
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location('picture_native', ROOT / 'scripts/check-xlsx-chart-colors.py')
native_module = importlib.util.module_from_spec(spec); spec.loader.exec_module(native_module)
native = native_module.native
fixture = native.fixture
REPORT = ROOT / 'docs/evidence/xlsx-picture-rotation.json'
FIXTURES = ROOT / 'packages/office/tests/fixtures'
A, D = '{' + fixture.A + '}', '{' + fixture.D + '}'


def inputs():
    return ([dict(rotation=angle, rotates=flag, flipH=False, flipV=False, crop=False) for angle in [0, 45, 90, -30] for flag in [None, True, False]]
            + [dict(rotation=45, rotates=False, flipH=h, flipV=v, crop=crop) for h, v in [(False, False), (True, False), (False, True), (True, True)] for crop in [False, True]])


def picture_png():
    def chunk(kind, data): return struct.pack('>I', len(data)) + kind + data + struct.pack('>I', zlib.crc32(kind + data))
    colors = [(255, 0, 0), (0, 255, 0), (0, 0, 255), (255, 255, 0)]
    pixels = b''.join(b'\0' + b''.join(bytes(colors[(y >= 32)*2 + (x >= 32)]) for x in range(64)) for y in range(64))
    return b'\x89PNG\r\n\x1a\n' + chunk(b'IHDR', struct.pack('>IIBBBBB', 64, 64, 8, 2, 0, 0, 0)) + chunk(b'IDAT', zlib.compress(pixels)) + chunk(b'IEND', b'')


def generate(cases):
    book = fixture.Book(); book.part('xl/media/quadrants.png', picture_png(), 'image/png')
    for index, case in enumerate(cases, 1):
        sheet = book.sheet('Picture' + str(index)); sheet.add('A1', 'Picture fill')
        flag = '' if case['rotates'] is None else ' rotWithShape="' + str(int(case['rotates'])) + '"'
        crop = '<a:srcRect l="10000" t="20000" r="25000" b="10000"/>' if case['crop'] else ''
        transform = fixture.attrs(dict(rot=case['rotation']*60000, flipH=int(case['flipH']), flipV=int(case['flipV'])))
        picture = ('<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="1" name="Quadrants" descr="Red green blue yellow quadrants"/><xdr:cNvPicPr/></xdr:nvPicPr>'
                   '<xdr:blipFill' + flag + '><a:blip r:embed="image"/>' + crop + '<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>'
                   '<xdr:spPr><a:xfrm' + transform + '><a:off x="0" y="0"/><a:ext cx="2286000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>')
        path = fixture.add_drawing(book, sheet, fixture.one_anchor(picture, 1, 3, 180, 90))
        book.relate(path, 'image', fixture.R + '/image', 'xl/media/quadrants.png')
    return fixture.zipped(book.finish())


def observe(data):
    with ZipFile(BytesIO(data)) as archive:
        result = []
        for index in range(1, len(inputs()) + 1):
            root = ET.fromstring(archive.read(f'xl/drawings/drawing{index}.xml'))
            picture = root.find('.//' + D + 'pic'); transform = picture.find(D + 'spPr/' + A + 'xfrm')
            fill = picture.find(D + 'blipFill'); crop = fill.find(A + 'srcRect')
            result.append(dict(fill=fill.attrib, transform=transform.attrib, crop=crop.attrib if crop is not None else {}))
        return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    font = next(item for item in json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts'] if item['file'] == 'LiberationSans-Regular.ttf')
    family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', 'Liberation Sans:style=Regular']).split('|')
    assert 'Liberation Sans' in family and Path(filename).name == font['file'] and sha256(Path(filename).read_bytes()).hexdigest() == font['sha256']
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'pictures.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        producer = (output / source.name).read_bytes()
        report = dict(scope=__doc__, office=native.BUILD, archiveSha256=native.ARCHIVE_HASH, fontManifest=[font], sourceSha256=sha256(data).hexdigest(),
                      specification='https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.spreadsheet.blipfill?view=openxml-3.0.1',
                      cases=[dict(**case, source=s, producer=p) for case, s, p in zip(cases, observe(data), observe(producer))],
                      limits=['Calc omits rotWithShape; its producer is validated with the properties actually saved, not reconstructed source flags.',
                              'Page-axis stretch uses the rotated frame bounding rectangle and its independently rotated clip. Browser geometry/pixel checks validate this interpretation; Excel/WPS remain unverified.'])
        if args.record:
            (FIXTURES / 'xlsx-picture-rotation.xlsx').write_bytes(data); (FIXTURES / 'xlsx-picture-rotation-producer.xlsx').write_bytes(producer)
            REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
        else:
            assert (FIXTURES / 'xlsx-picture-rotation.xlsx').read_bytes() == data
            assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(cases)} independent picture rotation sources and Calc properties')


if __name__ == '__main__': main()
