"""Measure the slant-dash-dot example in the original ECMA-376 standard."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import json
import subprocess

ROOT = Path(__file__).resolve().parent.parent
DIGEST = "9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129"
REPORT = ROOT / "docs/evidence/xlsx-border-spec.json"


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    archive_path = ROOT / ".cache/ecma376-part1-2016.zip"; assert sha256(archive_path.read_bytes()).hexdigest() == DIGEST
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); pdf = work / "standard.pdf"
        with ZipFile(archive_path) as archive: pdf.write_bytes(archive.read("Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf"))
        subprocess.run(["pdfimages", "-f", "2448", "-l", "2448", str(pdf), str(work / "example")], check=True, timeout=55)
        with (work / "example-003.ppm").open("rb") as file:
            assert file.readline().strip() == b"P6"; width, height = map(int, file.readline().split()); assert file.readline().strip() == b"255"
            pixels = file.read(); assert (width, height) == (86, 40) and len(pixels) == width * height * 3
    def black(x, y): return pixels[(y * width + x) * 3:(y * width + x) * 3 + 3] == b"\0\0\0"
    rows = ["".join("1" if black(x, y) else "0" for x in range(30, 66)) for y in [8, 9]]
    period = next(n for n in range(1, 19) if all(row[x] == row[x + n] for row in rows for x in range(len(row) - n)))
    assert all(not black(x, y) for y in [7, 10] for x in range(30, 66))
    report = {"source": "https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip", "archiveSha256": DIGEST,
              "section": "Part 1, fifth edition, 18.18.3; PDF page 2448 (printed 2438), fourth illustration.",
              "scope": "Functional stroke-mask measurement from the standard's 96-dpi illustration. Not an Excel/WPS runtime verification.",
              "imageRgbSha256": sha256(pixels).hexdigest(), "style": "slantDashDot", "dpi": 96, "widthPt": len(rows) * .75,
              "periodPt": period * .75, "rows": [row[:period] for row in rows]}
    if args.record: REPORT.write_text(json.dumps(report, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Standard border observations changed"
    print("Verified native slant-dash-dot standard illustration")


if __name__ == "__main__": main()
