"""Independent revision UIDs, classic GUIDs and modern compatibility identities."""
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
spec = importlib.util.spec_from_file_location('comment_uid_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-comment-uids.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-comment-uids.json'
XR = 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision'
MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
ROOT_ID = '{33333333-3333-4333-8333-333333333333}'
SECOND_ID = '{55555555-5555-4555-8555-555555555555}'
UID = '{DDDDDDDD-DDDD-4DDD-8DDD-DDDDDDDDDDDD}'
GUID = '{EEEEEEEE-EEEE-4EEE-8EEE-EEEEEEEEEEEE}'
MISMATCH = '{FFFFFFFF-FFFF-4FFF-8FFF-FFFFFFFFFFFF}'


def generate():
    book = native.fixture.Book(); sheet = book.sheet('CommentUIDs'); sheet.add('A1', 'UID source')
    native.fixture.add_modern(book, sheet); native.fixture.add_classic(book, sheet)
    path = 'xl/comments/comment1.xml'; text = book.parts[path]
    text = text.replace('<comments ', f'<comments xmlns:mc="{MC}" xmlns:revision="{XR}" mc:Ignorable="revision" ')
    text = text.replace('<comment ref="B2" authorId="0"', f'<comment ref="A1" authorId="0" guid="{GUID}" revision:uid="{UID}"')
    text = text.replace('</authors>', f'<author>tc={ROOT_ID}</author><author>tc={SECOND_ID}</author></authors>')
    caches = f'<comment ref="B2" authorId="1" revision:uid="{ROOT_ID}"><text><t>旧版兼容缓存</t></text></comment>'
    caches += f'<comment ref="C3" authorId="2" revision:uid="{MISMATCH}"><text><t>身份冲突的兼容缓存</t></text></comment>'
    book.parts[path] = text.replace('</commentList>', caches + '</commentList>')
    path = 'xl/drawings/comments.vml'; book.parts[path] = book.parts[path].replace('<x:Row>1</x:Row>', '<x:Row>0</x:Row>').replace('<x:Column>1</x:Column>', '<x:Column>0</x:Column>')
    return native.fixture.zipped(book.finish())


def comment_records(data):
    result = []
    with ZipFile(BytesIO(data)) as archive:
        for path in archive.namelist():
            if not path.endswith('.xml'): continue
            root = ET.fromstring(archive.read(path))
            if root.tag != '{' + native.fixture.S + '}comments': continue
            for comment in root.findall('{*}commentList/{*}comment'):
                text = comment.find('{*}text')
                result.append({'cell': comment.get('ref'), 'guid': comment.get('guid'), 'uid': comment.get('{' + XR + '}uid'), 'text': ''.join(text.itertext())})
    return result


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    data = generate()
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'uids.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        producer = (output / 'uids.xlsx').read_bytes()
    report = {'scope': __doc__, 'office': native.BUILD, 'archiveSha256': native.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(),
              'standard': 'MS-XLSX v20261016 sections 2.5.10 and 2.3.7.3; ISO/IEC 29500 comment guid remains separate.',
              'source': comment_records(data), 'producer': comment_records(producer),
              'limits': ['Source identities are opaque provenance, never authority.', 'UID mismatch is preserved for explicit thread recovery rather than moving or deleting a canonical modern thread.', 'Calc changes are recorded; Excel/WPS unverified.']}
    if args.record: FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else:
        assert FIXTURE.read_bytes() == data
        assert json.loads(REPORT.read_text()) == report
    print('Verified 3 independent revision UIDs and classic/modern provenance with pinned Calc')


if __name__ == '__main__': main()
