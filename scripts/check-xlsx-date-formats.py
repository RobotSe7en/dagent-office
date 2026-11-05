"""Record native date/time and locale observations without using SDK output.

Calc's historical calendar, whole-second truncation and locale vocabularies are
observations, not universal expected Excel results. SDK differences are tested
and documented separately; Excel/WPS have not been executed.
"""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
import argparse
import importlib.util
import json
import os

ROOT = Path(__file__).resolve().parent.parent
REPORT = ROOT / "docs/evidence/xlsx-date-format-semantics.json"
spec = importlib.util.spec_from_file_location("independent_dates", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
CASES = [(value, "yyyy-mm-dd ddd") for value in [0, 1, 59, 60, 61, 1462, 45351, 2957003]]
CASES += [(45351.5, "[$-" + locale + "]" + code) for locale in ["409", "804", "407", "40C", "419", "411", "412"]
          for code in ["ddd dddd mmm mmmm mmmmm", "d mmmm yyyy", "h:mm AM/PM", "h:mm A/P"]]
CASES += [(value, code) for value in [59.9 / 86400, 59.9996 / 86400, 1 - .0004 / 86400, 60 - .0004 / 86400, 1.999999999]
          for code in ["yyyy-mm-dd", "yyyy-mm-dd hh:mm:ss", "hh:mm", "hh:mm:ss.000", "[h]:mm:ss.00"]]
CASES += [(1.25, "[h]:mm"), (1.25, "[mm]:ss"), (1.25, "[ss].00"), (45351.5, '[$-409]dddd, mmmm d, yyyy "at" h:mm AM/PM')]


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    native.CASES = CASES; series = []
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for date1904 in [False, True]:
            system = "1904" if date1904 else "1900"; source = work / (system + ".xlsx"); source.write_bytes(native.generate(date1904))
            native.shared.run(["libreoffice", profile, "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
            observed = native.observe(source.with_suffix(".ods")); assert len(observed) == len(CASES)
            series.append({"dateSystem": system, "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                           "samples": [{"value": value, "code": code, "text": text} for (value, code), text in zip(CASES, observed)]})
    report = {"scope": __doc__.strip(), "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH,
              "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "cases": series}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native date/time observations changed"
    print(f"Verified {len(CASES) * 2} date/time and locale observations with pinned Calc")


if __name__ == "__main__": main()
