"""Observe native General section signs, affixes and system date/time formats.

Only ZIP/XML source generation and pinned Calc are used. Source syntax/cache
and SDK formatting remain separate. Locale data can use a different AM/PM
space; this is recorded rather than presented as an exact visual match.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os

ROOT = Path(__file__).resolve().parent.parent
REPORT = ROOT / "docs/evidence/xlsx-general-sections.json"
spec = importlib.util.spec_from_file_location("independent_general", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
CODES = ["General", "General;(General)", "[<0]General;0", 'General;"N"General', '"ID "General', 'General" kg"', "General General", "General%"]
CASES = [(value, code) for value in [-12.5, 12.5] for code in CODES]
EXPECTED = ["-12.5", "(12.5)", "12.5", "N12.5", "-ID 12.5", "-12.5 kg", "-12.5 12.5", "-1250%",
            "12.5", "12.5", "13", "12.5", "ID 12.5", "12.5 kg", "12.5 12.5", "1250%"]
CASES += [(45351.5, code) for code in ["[$-F800]yyyy-mm-dd", '[$-F800]"ignored"yyyy-mm-dd', "[$-F400]h:mm"]]
EXPECTED += ["Thursday, February 29, 2024", "Thursday, February 29, 2024", "12:00:00 PM"]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    native.CASES = CASES
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "general.xlsx"; source.write_bytes(native.generate(False))
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        observed = native.observe(source.with_suffix(".ods")); assert observed == EXPECTED, observed
        report = {"scope": __doc__.strip(), "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "dateSystem": "1900",
                  "sourceSha256": sha256(source.read_bytes()).hexdigest(), "hostLocale": "en-US",
                  "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"},
                  "samples": [{"value": value, "code": code, "text": text} for (value, code), text in zip(CASES, observed)]}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native General section observations changed"
    print(f"Verified {len(CASES)} independent native General/system-format results")


if __name__ == "__main__": main()
