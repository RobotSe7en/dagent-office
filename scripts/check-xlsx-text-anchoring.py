"""Observe independent DrawingML block anchoring with pinned Calc and fonts.

The source is ZIP/XML, independent of the SDK. Producer geometry is evidence,
not an Excel/WPS claim. Normal checks never replace the recorded observations.
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
import re

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("anchoring_text_probe", ROOT / "scripts/check-xlsx-drawing-text.py")
source = importlib.util.module_from_spec(spec); spec.loader.exec_module(source)
native = source.native
REPORT = ROOT / "docs/evidence/xlsx-text-anchoring.json"
FIXTURE = ROOT / "packages/office/tests/fixtures/xlsx-text-anchoring.xlsx"


def inputs():
    return [{"fontPt": 14, "endPt": 14, "paragraphs": ["Alpha", "Beta"], "insets": "explicit-default", "listStyle": "none",
             "anchor": anchor, "anchorCenter": centered, "alignments": alignment}
            for alignment in [["l", "l"], ["ctr", "ctr"], ["r", "r"], ["l", "r"]]
            for anchor in ["t", "ctr", "b"] for centered in [False, True]]


def generate(cases):
    with ZipFile(BytesIO(source.generate(cases))) as archive:
        parts = {name: archive.read(name) for name in archive.namelist()}
    for index, case in enumerate(cases, 1):
        path = f"xl/drawings/drawing{index}.xml"
        xml = parts[path].decode()
        xml = xml.replace('<a:bodyPr ', f'<a:bodyPr anchor="{case["anchor"]}" anchorCtr="{int(case["anchorCenter"])}" ')
        alignments = iter(case["alignments"])
        xml = re.sub(r"<a:pPr/>", lambda _: f'<a:pPr algn="{next(alignments)}"/>', xml)
        parts[path] = xml
    return native.fixture.zipped(parts)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = native.shared.environment(); assert environment["libreOffice"] == native.BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]; assert environment["fonts"] == fonts
    cache = ROOT / ".cache/xlsx-text-anchoring"; cache.mkdir(exist_ok=True)
    cases = inputs(); data = generate(cases)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); output = work / "out"; output.mkdir(); path = work / "anchoring.xlsx"; path.write_bytes(data)
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["pdf", "xlsx:Calc MS Excel 2007 XML"]:
            native.shared.run(["libreoffice", profile, "--headless", "--convert-to", kind, "--outdir", str(output), str(path)])
        observations = source.observe(output / "anchoring.pdf", output / "anchoring.xlsx", cases, work)
    report = {"scope": "Pinned Calc DrawingML anchor/anchorCtr with left, center, right and mixed paragraph alignment. Excel/WPS unverified.",
              "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(data).hexdigest(),
              "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()},
              "cases": observations}
    if args.record:
        FIXTURE.write_bytes(data); REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else:
        assert FIXTURE.read_bytes() == data, "Independent text anchoring source changed"
        assert json.loads(REPORT.read_text()) == report, "Native text anchoring evidence changed"
    print(f"Verified {len(cases)} independent text-anchoring probes with pinned Calc")


if __name__ == "__main__": main()
