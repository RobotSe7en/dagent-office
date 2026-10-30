"""Render public DOCX fixtures with LibreOffice and record an independent baseline.
Default mode writes only .cache and verifies the recorded environment and PNGs.
Use --record deliberately to refresh tracked references and producer DOCX inputs.
"""
from pathlib import Path
import argparse
import hashlib
import json
import shutil
import subprocess
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
parser = argparse.ArgumentParser()
parser.add_argument('--record', action='store_true')
args = parser.parse_args()
samples = ROOT / 'examples/docx'
cache = ROOT / '.cache/docx-reference'
producer = cache / 'producer'
cache.mkdir(parents=True, exist_ok=True)
producer.mkdir(exist_ok=True)

def run(command):
    result = subprocess.run(command, cwd=ROOT, capture_output=True, text=True, timeout=55)
    if result.returncode:
        raise RuntimeError(result.stderr or result.stdout or str(command))
    return result.stdout.strip()

def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

fonts = []
for family in ['Noto Sans CJK SC', 'Noto Serif CJK SC', 'Liberation Serif']:
    path = Path(run(['fc-match', '-f', '%{file}', family]))
    actual = run(['fc-match', '-f', '%{family}', family])
    if family not in actual:
        raise RuntimeError('Required reference font unavailable: ' + family)
    fonts.append({'family': family, 'file': path.name, 'sha256': digest(path)})

shells = sorted((ROOT / '.playwright').glob('chromium_headless_shell-*/chrome-headless-shell-linux64/chrome-headless-shell'))
if not shells:
    raise RuntimeError('Install the pinned browser with pnpm browser:install first.')
environment = {'libreOffice': run(['libreoffice', '--version']), 'chromium': run([str(shells[-1]), '--version']), 'dpi': 96, 'fonts': fonts}
inputs = sorted(samples.glob('*.docx'))
profile = '-env:UserInstallation=' + (cache/'profile').as_uri()
run(['libreoffice', profile, '--headless', '--convert-to', 'pdf', '--outdir', str(cache), *map(str,inputs)])
run(['libreoffice', profile, '--headless', '--convert-to', 'docx', '--outdir', str(producer), *map(str,inputs)])
records = []
for source in inputs:
    pdf = cache/(source.stem+'.pdf')
    if not pdf.exists():
        raise RuntimeError('Reference conversion did not produce ' + pdf.name)
    prefix = cache/source.stem
    run(['pdftoppm', '-r', '96', '-png', str(pdf), str(prefix)])
    words = run(['pdftotext', '-bbox', str(pdf), '-'])
    root = ET.fromstring(words)
    page_count = len(root.findall('.//{http://www.w3.org/1999/xhtml}page'))
    images = sorted(cache.glob(source.stem+'-[0-9]*.png'))
    if not images or len(images) != page_count:
        raise RuntimeError('Reference page count mismatch: ' + source.name)
    record = {'file': source.name, 'sha256': digest(source), 'pages': page_count,
              'images': [{'file': image.name, 'sha256': digest(image)} for image in images],
              'producerFile': 'producer/'+source.name}
    records.append(record)
    if args.record:
        (samples/'reference').mkdir(exist_ok=True)
        (samples/'producer').mkdir(exist_ok=True)
        for image in images:
            shutil.copyfile(image, samples/'reference'/image.name)
        shutil.copyfile(producer/source.name, samples/'producer'/source.name)
record = {'environment':environment, 'fixtures':records,
          'limits':['Continuous preview is not automatic pagination.', 'Fields are checked against source caches, not LibreOffice updates.', 'Insertion/deletion marks are checked structurally; PDF review display can differ.', 'Fixture 09 intentionally requires font substitution and includes a missing CJK glyph.', 'Word and WPS acceptance is not recorded by this local reference.']}
path = samples/'reference.json'
if args.record:
    path.write_text(json.dumps(record, ensure_ascii=False, indent=2)+'\n')
else:
    expected = json.loads(path.read_text())
    if record != expected:
        raise RuntimeError('Reference environment, source or rendered pages differ; inspect changes before --record.')
print('Verified' if not args.record else 'Recorded', len(records), 'independent DOCX reference cases; fonts and renderer versions captured.')
