"""Independent classic note sources with VML and commentPr boxes.

The selected display source is VML. Both mapped native sources and property
differences are retained; Calc resave losses are recorded independently.
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
spec = importlib.util.spec_from_file_location('note_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
S = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
XDR = '{http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing}'
V = '{urn:schemas-microsoft-com:vml}'
X = '{urn:schemas-microsoft-com:office:excel}'
BASE = ROOT / 'examples/xlsx/09-classic-notes.xlsx'
FIXTURES = ROOT / 'packages/office/tests/fixtures'
REPORT = ROOT / 'docs/evidence/xlsx-dual-note-boxes.json'
CASES = ['conflict', 'same-anchor', 'visible', 'commentPr-only']


def generate(case):
    with ZipFile(BASE) as archive: parts = {name: archive.read(name) for name in archive.namelist()}
    part = 'xl/comments/comment1.xml'; root = ET.fromstring(parts[part]); comment = root.find(S + 'commentList/' + S + 'comment')
    properties = ET.SubElement(comment, S + 'commentPr', {'print': '0', 'autoFill': '0', 'altText': '备用备注框😀', 'textHAlign': 'center', 'textVAlign': 'bottom'})
    anchor = ET.SubElement(properties, S + 'anchor', {'moveWithCells': '1', 'sizeWithCells': '1' if case == 'same-anchor' else '0'})
    for name, values in [('from', [2, 95250, 1, 19050] if case == 'same-anchor' else [3, 0, 2, 0]), ('to', [5, 190500, 5, 38100] if case == 'same-anchor' else [7, 127000, 8, 25400])]:
        marker = ET.SubElement(anchor, XDR + name)
        for key, value in zip(['col', 'colOff', 'row', 'rowOff'], values): ET.SubElement(marker, XDR + key).text = str(value)
    parts[part] = ET.tostring(root, encoding='unicode')
    if case == 'visible':
        part = 'xl/drawings/comments.vml'; root = ET.fromstring(parts[part]); shape = root.find(V + 'shape')
        shape.set('style', shape.get('style').replace('visibility:hidden', 'visibility:visible')); ET.SubElement(shape.find(X + 'ClientData'), X + 'Visible')
        parts[part] = ET.tostring(root, encoding='unicode')
    if case == 'commentPr-only':
        part = 'xl/worksheets/sheet1.xml'; root = ET.fromstring(parts[part]); root.remove(root.find(S + 'legacyDrawing')); parts[part] = ET.tostring(root, encoding='unicode')
        part = 'xl/worksheets/_rels/sheet1.xml.rels'; root = ET.fromstring(parts[part])
        for node in list(root):
            if node.get('Type').endswith('/vmlDrawing'): root.remove(node)
        parts[part] = ET.tostring(root, encoding='unicode'); del parts['xl/drawings/comments.vml']
        part = '[Content_Types].xml'; root = ET.fromstring(parts[part])
        for node in list(root):
            if node.get('PartName') == '/xl/drawings/comments.vml': root.remove(node)
        parts[part] = ET.tostring(root, encoding='unicode')
    return native.fixture.zipped(parts)


def observe(data):
    with ZipFile(BytesIO(data)) as archive:
        comments = [ET.fromstring(archive.read(name)) for name in archive.namelist() if name.startswith('xl/comments') and name.endswith('.xml')]
        notes = [node for root in comments for node in root.findall(S + 'commentList/' + S + 'comment')]
        shapes = [node for name in archive.namelist() if name.endswith('.vml') for node in ET.fromstring(archive.read(name)).findall(V + 'shape')]
        return {'commentCells': [note.get('ref') for note in notes], 'commentPr': [ET.tostring(note.find(S + 'commentPr'), encoding='unicode') if note.find(S + 'commentPr') is not None else None for note in notes],
                'vml': [{'anchor': shape.find(X + 'ClientData/' + X + 'Anchor').text, 'visible': shape.find(X + 'ClientData/' + X + 'Visible') is not None} for shape in shapes]}


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']; os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    baseline = json.loads((ROOT / 'examples/xlsx/reference.json').read_text())['environment']['fonts']; fonts = []
    for pattern in ['Liberation Sans:style=Regular', 'Liberation Sans:style=Bold', 'Noto Sans CJK SC:style=Regular', 'Noto Sans CJK SC:style=Bold', 'Noto Color Emoji']:
        family, filename = native.shared.run(['fc-match', '-f', '%{family}|%{file}', pattern]).split('|')
        assert pattern.split(':')[0] in family, 'Required note reference family did not resolve'
        expected = next(font for font in baseline if font['file'] == Path(filename).name)
        assert sha256(Path(filename).read_bytes()).hexdigest() == expected['sha256'], 'Required note reference font differs'
        if expected not in fonts: fonts.append(expected)
    records = []
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary)
        for case in CASES:
            data = generate(case); name = 'xlsx-dual-note-boxes-' + case; source = work / (name + '.xlsx'); source.write_bytes(data); output = work / ('out-' + case); output.mkdir()
            native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / ('profile-' + case)).as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
            producer = (output / source.name).read_bytes()
            records.append(dict(case=case, sourceSha256=sha256(data).hexdigest(), source=observe(data), producer=observe(producer)))
            if args.record:
                (FIXTURES / (name + '.xlsx')).write_bytes(data); (FIXTURES / (name + '-producer.xlsx')).write_bytes(producer)
            else: assert (FIXTURES / (name + '.xlsx')).read_bytes() == data
    report = dict(scope=__doc__, baseSha256=sha256(BASE.read_bytes()).hexdigest(), archiveSha256=native.ARCHIVE_HASH, office=native.BUILD, fontManifest=fonts, cases=records,
                  limits=['Office implementation notes say Excel rejects commentPr; the user selected VML priority, not an inferred Excel precedence.', 'Calc resave may discard commentPr or replace its geometry; source retention and producer reopen are separate checks.', 'No Excel/WPS machine acceptance or complete classic-box appearance comparison is claimed.'])
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else: assert json.loads(REPORT.read_text()) == report
    print(f'Verified {len(CASES)} independent dual classic-note sources and Calc resaves')


if __name__ == '__main__': main()
