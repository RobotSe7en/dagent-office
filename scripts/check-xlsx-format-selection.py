"""Record native sign, section and text behavior without SDK-generated files."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("native_format_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-format-selection.json"
CODES = ['0.00', '0.00;(0.00);"zero"', '[>10]"A"0;[<0]"B"0;"C"0', '[>10]"A"0;"B"0',
         '0;@', '0;"neg"0;@', '0;0;"zero";"ID "@', '@', '"constant"']
VALUES = [-1, 0, 5, 20, "001😀"]
EXPECTED = [["-1.00", "0.00", "5.00", "20.00", "001😀"], ["(1.00)", "zero", "5.00", "20.00", "001😀"],
            ["B1", "C0", "C5", "A20", "001😀"], ["B1", "B0", "B5", "A20", "001😀"],
            ["-1", "0", "5", "20", "001😀"], ["neg1", "0", "5", "20", "001😀"],
            ["1", "zero", "5", "20", "ID 001😀"], ["-1", "0", "5", "20", "001😀"],
            ["constant", "constant", "constant", "constant", "001😀"]]
native.CASES = [(value, code) for code in CODES for value in VALUES] + [(-.001, '0.00'), (-.001, '0.00;(0.00)')]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "selection.xlsx"; source.write_bytes(native.generate(False))
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        observed = native.observe(source.with_suffix(".ods"))
        assert observed == [text for group in EXPECTED for text in group] + ["0.00", "(0.00)"], observed
        report = {"scope": "Selected native Calc sign, section and text results. No SDK General fitting, date display, Excel/WPS or complete custom-format acceptance claim.",
                  "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                  "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"},
                  "samples": [{"value": value, "code": code, "text": text} for (value, code), text in zip(native.CASES, observed)]}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native format section results changed"
    print(f"Verified {len(native.CASES)} native number-format sign/section/text results")


if __name__ == "__main__": main()
