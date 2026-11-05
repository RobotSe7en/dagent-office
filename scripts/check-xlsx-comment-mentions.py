"""Independent modern mentions and conflicting cell provenance; no notifications."""
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
spec = importlib.util.spec_from_file_location('comment_mentions_native', ROOT / 'scripts/check-xlsx-number-formats.py')
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
FIXTURE = ROOT / 'packages/office/tests/fixtures/xlsx-comment-mentions.xlsx'
REPORT = ROOT / 'docs/evidence/xlsx-comment-mentions.json'
TC = native.fixture.TC
PERSON = '{22222222-2222-4222-8222-222222222222}'


def mention(identifier, start, length):
    return f'<mentions><mention mentionpersonId="{PERSON}" mentionId="{identifier}" startIndex="{start}" length="{length}"/></mentions>'


def generate():
    book = native.fixture.Book(); sheet = book.sheet('Mentions'); sheet.add('A1', 'Independent mentions'); native.fixture.add_modern(book, sheet)
    person = '{DDDDDDDD-DDDD-4DDD-8DDD-DDDDDDDDDDDD}'
    book.parts['xl/persons/person.xml'] = book.parts['xl/persons/person.xml'].replace('</personList>', f'<person displayName="只被提及" id="{person}" userId="opaque-mentioned-only" providerId="AD"/></personList>')
    path = 'xl/threadedComments/threadedComment1.xml'; text = book.parts[path]
    text = text.replace('</text>', '</text>' + mention('{66666666-6666-4666-8666-666666666666}', 5, 5), 1)
    text = text.replace('ref="B2" dT="2026-11-04T08:01:00Z"', 'ref="B3" dT="2026-11-04T08:01:00Z"').replace('ref="C3"', 'ref="B2"')
    text = text.replace('<text>已复核并补充。</text>', '<text>已复核并补充。</text>' + mention('{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}', 0, 3))
    extra = f'<threadedComment ref="D4" dT="2026-11-04T09:00:00Z" personId="{PERSON}" id="{{88888888-8888-4888-8888-888888888888}}"><text>@复核人😀禰󠄀</text>' + mention('{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}', 5, 1) + '</threadedComment>'
    extra += f'<threadedComment dT="2026-11-04T09:01:00Z" personId="{PERSON}" id="{{99999999-9999-4999-8999-999999999999}}"><text>缺少位置的线程</text></threadedComment>'
    extra += f'<threadedComment ref="E5" dT="2026-11-04T09:02:00Z" personId="{PERSON}" id="{{77777777-7777-4777-8777-777777777777}}"><text>@只被提及 请看😀</text>' + mention('{CCCCCCCC-CCCC-4CCC-8CCC-CCCCCCCCCCCC}', 0, 5).replace(PERSON, person) + '</threadedComment>'
    book.parts[path] = text.replace('</ThreadedComments>', extra + '</ThreadedComments>')
    return native.fixture.zipped(book.finish())


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument('--record', action='store_true'); args = parser.parse_args()
    assert sha256((ROOT / '.cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz').read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ['PATH'] = str(ROOT / '.cache/acceptance-office/bin') + os.pathsep + os.environ['PATH']
    os.environ['LD_LIBRARY_PATH'] = str(ROOT / '.cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu')
    assert native.shared.run(['libreoffice', '--version']) == native.BUILD
    data = generate()
    with TemporaryDirectory(dir=ROOT / '.cache') as temporary:
        work = Path(temporary); source = work / 'mentions.xlsx'; source.write_bytes(data); output = work / 'out'; output.mkdir()
        native.shared.run(['libreoffice', '-env:UserInstallation=' + (work / 'profile').as_uri(), '--headless', '--convert-to', 'xlsx:Calc MS Excel 2007 XML', '--outdir', str(output), str(source)])
        with ZipFile(output / 'mentions.xlsx') as archive: retained = [name for name in archive.namelist() if 'threadedcomment' in name.lower() or '/persons/' in name.lower()]
    # External omission is recorded; the SDK must retain all source messages.
    assert retained == []
    with ZipFile(BytesIO(data)) as archive:
        people = [person.attrib for person in ET.fromstring(archive.read('xl/persons/person.xml'))]
        root = ET.fromstring(archive.read('xl/threadedComments/threadedComment1.xml'))
        messages = [{'id': node.get('id'), 'cell': node.get('ref'), 'parentId': node.get('parentId'), 'text': node.findtext('{' + TC + '}text'),
                     'mentions': [mention.attrib for mention in node.findall('{' + TC + '}mentions/{' + TC + '}mention')]} for node in root]
    report = {'scope': __doc__, 'office': native.BUILD, 'archiveSha256': native.ARCHIVE_HASH, 'sourceSha256': sha256(data).hexdigest(),
              'standard': 'https://learn.microsoft.com/en-us/openspecs/office_standards/ms-xlsx/b03ed619-e307-4d3e-9c67-b68612274128',
              'people': people, 'messages': messages, 'producerRetainedThreadParts': retained,
              'limits': ['Calc removes the modern source; this is an external loss, not fidelity acceptance.', 'Native offsets are retained as UTF-16 units. Splitting a surrogate pair or variation sequence requires reanchoring.', 'Excel/WPS unverified.']}
    if args.record: FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + '\n')
    else:
        assert FIXTURE.read_bytes() == data
        assert json.loads(REPORT.read_text()) == report
    print('Verified 6 independent modern messages and 4 mentions; Calc omission recorded')


if __name__ == '__main__': main()
