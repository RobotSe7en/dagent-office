"""Measure binary fill patterns from the original ECMA-376 illustrations."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import json
import subprocess

ROOT = Path(__file__).resolve().parent.parent
ARCHIVE = ROOT / ".cache/ecma376-part1-2016.zip"
DIGEST = "9d0bcad9cf06054785b03762fcfadbf6bab7e54a5f9d69434e34b7fd464d4129"
REPORT = ROOT / "docs/evidence/xlsx-fill-pattern-spec.json"
NAMES = ["darkDown", "darkGray", "darkGrid", "darkHorizontal", "darkTrellis", "darkUp", "darkVertical", "gray0625", "gray125", "lightDown", "lightGray", "lightGrid", "lightHorizontal", "lightTrellis", "lightUp", "mediumGray"]


def observe(path, name, index):
    with path.open("rb") as file:
        assert file.readline().strip() == b"P6"
        width, height = map(int, file.readline().split()); assert file.readline().strip() == b"255"
        pixels = file.read(); assert len(pixels) == width * height * 3
    def rgb(x, y): return tuple(pixels[(y * width + x) * 3:(y * width + x) * 3 + 3])
    yellow = [(x, y) for y in range(height) for x in range(width) if rgb(x, y) == (255, 255, 0)]
    x0 = min(x for x, y in yellow) + 16; y0 = min(y for x, y in yellow) + 8
    assert all(rgb(x0 + x, y0 + y) in [(0, 0, 0), (255, 255, 0)] for y in range(64) for x in range(64))
    samples = [[int(rgb(x0 + x, y0 + y) == (0, 0, 0)) for x in range(64)] for y in range(64)]
    px = next(p for p in range(1, 33) if all(row[x] == row[x + p] for row in samples for x in range(64 - p)))
    py = next(p for p in range(1, 33) if all(samples[y][x] == samples[y + p][x] for y in range(64 - p) for x in range(64)))
    # The printed examples magnify each logical dot into a 4x4 pixel square.
    assert all(samples[y][x] == samples[y // 4 * 4][x // 4 * 4] for y in range(64) for x in range(64))
    rows = ["".join(str(samples[y][x]) for x in range(0, px, 4)) for y in range(0, py, 4)]
    return {"pattern": name, "imageIndex": index, "imageRgbSha256": sha256(pixels).hexdigest(), "illustrationScale": 4,
            "rows": rows, "foregroundFraction": sum(row.count("1") for row in rows) / sum(map(len, rows))}


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256(ARCHIVE.read_bytes()).hexdigest() == DIGEST, "Original ECMA archive is missing or differs"
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); pdf = work / "standard.pdf"
        with ZipFile(ARCHIVE) as archive: pdf.write_bytes(archive.read("Ecma Office Open XML Part 1 - Fundamentals And Markup Language Reference.pdf"))
        subprocess.run(["pdfimages", "-f", "2481", "-l", "2484", str(pdf), str(work / "example")], check=True, timeout=55)
        cases = [observe(work / f"example-{index:03}.ppm", name, index) for index, name in enumerate(NAMES)]
    horizontal = next(case for case in cases if case["pattern"] == "lightHorizontal")
    cases.append({"pattern": "lightVertical", "basis": "Vertical counterpart of lightHorizontal; the standard names this direction but provides no separate illustration.",
                  "rows": ["".join(row[0] for row in horizontal["rows"])], "foregroundFraction": horizontal["foregroundFraction"]})
    report = {"source": "https://ecma-international.org/wp-content/uploads/ECMA-376-1_5th_edition_december_2016.zip", "archiveSha256": DIGEST,
              "section": "Part 1, fifth edition, section 18.18.55; PDF pages 2481–2484 (printed 2471–2474).",
              "scope": "Binary pattern geometry and coverage, up to repeating phase. Functional measurements from standard illustrations; no Office conversion library or SDK output.", "cases": cases}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Standard pattern measurements changed"
    print(f"Verified {len(cases)} standard fill patterns")


if __name__ == "__main__": main()
