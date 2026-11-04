"""Record independent native HLS tint results and their 8-bit rounding limits."""
from hashlib import sha256
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import colorsys
import importlib.util
import json
import os
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("native_tint_probe", ROOT / "scripts/check-xlsx-number-formats.py")
native = importlib.util.module_from_spec(spec); spec.loader.exec_module(native)
REPORT = ROOT / "docs/evidence/xlsx-color-tint-semantics.json"
STYLE = "urn:oasis:names:tc:opendocument:xmlns:style:1.0"
FO = "urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
COLORS = ["000000", "FFFFFF", "4472C4", "ED7D31", "70AD47", "FF0000", "00FF00", "0000FF", "123456", "888888", "7F8081", "C1AB29"]
CASES = [(color, tint) for color in COLORS for tint in [-1, -.8, -.5, -.2, 0, .2, .4, .8, 1]]


def generate():
    book = native.fixture.Book(); sheet = book.sheet("Colors"); namespace = "{" + native.fixture.S + "}"
    styles = ET.fromstring(book.parts["xl/styles.xml"]); fills = styles.find(namespace + "fills")
    for child in list(fills)[2:]: fills.remove(child)
    fills.set("count", str(len(CASES) + 2))
    cells = styles.find(namespace + "cellXfs")
    for child in list(cells)[1:]: cells.remove(child)
    cells.set("count", str(len(CASES) + 1))
    for index, (color, tint) in enumerate(CASES):
        pattern = ET.SubElement(ET.SubElement(fills, namespace + "fill"), namespace + "patternFill", {"patternType": "solid"})
        ET.SubElement(pattern, namespace + "fgColor", {"rgb": "FF" + color, "tint": str(tint)})
        ET.SubElement(cells, namespace + "xf", {"fontId": "0", "fillId": str(index + 2), "borderId": "0", "numFmtId": "0", "xfId": "0", "applyFill": "1"})
        sheet.add("A" + str(index + 1), str(index), style=index + 1)
    book.parts["xl/styles.xml"] = ET.tostring(styles, encoding="unicode")
    return native.fixture.zipped(book.finish())


def observe(path):
    with ZipFile(path) as archive: roots = [ET.fromstring(archive.read(part)) for part in ["styles.xml", "content.xml"]]
    definitions = {node.get(f"{{{STYLE}}}name"): node for root in roots for node in root.iter(f"{{{STYLE}}}style")}
    column_default = roots[1].find(f".//{{{native.T}}}table-column").get(f"{{{native.T}}}default-cell-style-name")
    cells = [cell for cell in roots[1].iter(f"{{{native.T}}}table-cell") if cell.get(f"{{{native.O}}}value-type")]
    assert len(cells) == len(CASES)
    colors = []
    for index, cell in enumerate(cells):
        assert "".join(cell.itertext()) == str(index)
        node = definitions[cell.get(f"{{{native.T}}}style-name") or column_default]
        colors.append(node.find(f"{{{STYLE}}}table-cell-properties").get(f"{{{FO}}}background-color"))
    return colors


def channels(color): return [int(color[offset:offset + 2], 16) for offset in [1, 3, 5]]


def hls_color(color, tint):
    hue, light, saturation = colorsys.rgb_to_hls(*(value / 255 for value in channels("#" + color)))
    light = light * (1 + tint) if tint < 0 else light * (1 - tint) + tint
    return "#" + "".join(f"{int(value * 255 + .5):02x}" for value in colorsys.hls_to_rgb(hue, light, saturation))


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true"); args = parser.parse_args()
    assert sha256((ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz").read_bytes()).hexdigest() == native.ARCHIVE_HASH
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    os.environ["LD_LIBRARY_PATH"] = str(ROOT / ".cache/acceptance-office-deps/installed/usr/lib/x86_64-linux-gnu")
    os.environ.update({"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"})
    assert native.shared.run(["libreoffice", "--version"]) == native.BUILD
    with TemporaryDirectory(dir=ROOT / ".cache") as temporary:
        work = Path(temporary); source = work / "colors.xlsx"; source.write_bytes(generate())
        native.shared.run(["libreoffice", "-env:UserInstallation=" + (work / "profile").as_uri(), "--headless", "--convert-to", "ods", "--outdir", str(work), str(source)])
        samples = []
        for (color, tint), observed in zip(CASES, observe(source.with_suffix(".ods"))):
            continuous = hls_color(color, tint)
            delta = max(abs(a - b) for a, b in zip(channels(observed), channels(continuous)))
            assert delta <= 1, (color, tint, observed, continuous)
            samples.append({"rgb": "#" + color.lower(), "tint": tint, "native": observed, "continuousHls": continuous, "maximumChannelDelta": delta})
        report = {"scope": "Native Calc SpreadsheetML HLS tint observations. Continuous HLS conversion differs at some 8-bit rounding ties by at most one channel unit; it is not a byte-exact native-color or Excel/WPS visual acceptance claim.",
                  "office": native.BUILD, "archiveSha256": native.ARCHIVE_HASH, "sourceSha256": sha256(source.read_bytes()).hexdigest(),
                  "localeEnvironment": {"LC_ALL": "C.UTF-8", "LANG": "C.UTF-8", "TZ": "UTC"}, "samples": samples}
    if args.record: REPORT.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n")
    else: assert json.loads(REPORT.read_text()) == report, "Native color tint observations changed"
    print(f"Verified {len(samples)} native HLS tint results with maximum one-channel-unit rounding difference")


if __name__ == "__main__": main()
