"""Probe classic note anchor precedence with independent ZIP/XML and pinned Calc.

This is external acceptance evidence only, not an SDK conversion dependency.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, ZipInfo, ZIP_DEFLATED
import argparse
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
BUILD = "LibreOffice 25.8.1.1 54047653041915e595ad4e45cccea684809c77b5"
ARCHIVE_HASH = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
REPORT = ROOT / "docs/evidence/xlsx-note-semantics.json"
V, X = "urn:schemas-microsoft-com:vml", "urn:schemas-microsoft-com:office:excel"
spec = importlib.util.spec_from_file_location("office_reference_tools", ROOT / "scripts/pptx-reference.py")
shared = importlib.util.module_from_spec(spec); spec.loader.exec_module(shared)


def generate(destination, kind):
    changes = {"base": {}, "different-css": {
        "margin-left:80pt;margin-top:30pt;width:150pt;height:75pt": "margin-left:600pt;margin-top:700pt;width:20pt;height:10pt"},
        "different-anchor": {"2, 10, 1, 2, 5, 20, 5, 4": "1, 40, 2, 20, 4, 80, 8, 40"}}
    with ZipFile(ROOT / "examples/xlsx/09-classic-notes.xlsx") as source, ZipFile(destination, "w", ZIP_DEFLATED) as output:
        for path in source.namelist():
            data = source.read(path)
            if path.endswith(".vml"):
                text = data.decode()
                for before, after in changes[kind].items():
                    assert before in text; text = text.replace(before, after)
                data = text.encode()
            entry = ZipInfo(path, (1980, 1, 1, 0, 0, 0)); entry.compress_type = ZIP_DEFLATED; output.writestr(entry, data)


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    environment = shared.environment(); assert environment["libreOffice"] == BUILD
    fonts = json.loads((ROOT / "examples/xlsx/reference.json").read_text())["environment"]["fonts"]
    assert environment["fonts"] == fonts, "The fixed XLSX font environment differs"
    font_digest = sha256(json.dumps(fonts, sort_keys=True, separators=(",", ":")).encode()).hexdigest()
    cases = []; cache = ROOT / ".cache/xlsx-note-semantics"; cache.mkdir(parents=True, exist_ok=True)
    with TemporaryDirectory(dir=cache) as temporary:
        work = Path(temporary); (work / "out").mkdir(); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for kind in ["base", "different-css", "different-anchor"]:
            source = work / (kind + ".xlsx"); generate(source, kind)
            shared.run(["libreoffice", profile, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(work / "out"), str(source)])
            with ZipFile(work / "out" / source.name) as archive:
                parts = [p for p in archive.namelist() if p.endswith(".vml")]; assert len(parts) == 1
                root = ET.fromstring(archive.read(parts[0])); shapes = root.findall(f"{{{V}}}shape"); assert len(shapes) == 1
                shape = shapes[0]; style = dict(item.split(":", 1) for item in shape.get("style").split(";") if item)
                observed = {"geometryPt": {key: float(style[key].removesuffix("pt")) for key in ["margin-left", "margin-top", "width", "height"]},
                            "anchor": shape.find(f".//{{{X}}}Anchor").text, "visibility": style["visibility"]}
            cases.append({"case": kind, "sourceSha256": sha256(source.read_bytes()).hexdigest(), "observed": observed})
    assert cases[0]["observed"] == cases[1]["observed"], "Calc unexpectedly prioritized VML CSS coordinates over its cell anchor"
    assert cases[2]["observed"]["geometryPt"] != cases[0]["observed"]["geometryPt"], "Changing the cell anchor did not change note geometry"
    report = {"scope": "Classic note VML anchor precedence only; Excel/WPS and SDK visual acceptance are unverified.",
              "office": BUILD, "archiveSha256": ARCHIVE_HASH, "fontManifest": {"source": "examples/xlsx/reference.json#environment.fonts", "sha256": font_digest}, "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native note geometry evidence changed"
    print(f"Verified {len(cases)} independent classic-note anchor probes with pinned Calc")


if __name__ == "__main__": main()
