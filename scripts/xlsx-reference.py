"""Record or verify pinned Calc resaves and fixed-font XLSX PDF/PNG references.

Only --record writes reference baselines. Source and producer feature counts are
reported separately, including unsupported modern-thread loss in external tools.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import importlib.util
import json
import os
import re
import shutil
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "examples/xlsx"
CACHE = ROOT / ".cache/xlsx-reference"
spec = importlib.util.spec_from_file_location("office_reference_tools", ROOT / "scripts/pptx-reference.py")
shared = importlib.util.module_from_spec(spec); spec.loader.exec_module(shared)
run, digest = shared.run, shared.digest
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
D = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"
TC = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"


def features(path):
    counts = {name: 0 for name in ["sheets", "cells", "formulas", "tables", "conditionalRules", "validations", "charts", "images", "shapes", "classicNotes", "threads", "replies", "resolvedThreads"]}
    tags = {"sheet": "sheets", "c": "cells", "f": "formulas", "table": "tables", "cfRule": "conditionalRules", "dataValidation": "validations", "comment": "classicNotes"}
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith(".xml"): continue
            root = ET.fromstring(archive.read(name))
            for node in root.iter():
                if node.tag.startswith("{" + S + "}"):
                    key = tags.get(node.tag.split("}")[1])
                    if key: counts[key] += 1
                if node.tag == "{" + D + "}pic": counts["images"] += 1
                if node.tag == "{" + D + "}sp": counts["shapes"] += 1
                if node.tag == "{http://schemas.openxmlformats.org/drawingml/2006/chart}chart": counts["charts"] += 1
                if node.tag == "{" + TC + "}threadedComment":
                    counts["replies" if node.get("parentId") else "threads"] += 1
                    if not node.get("parentId") and node.get("done") in ["1", "true"]: counts["resolvedThreads"] += 1
    return counts


def verify_saved(baseline):
    for item in baseline["fixtures"]:
        files = [(item["file"], item["sha256"]), (item["producerFile"], item["producerSha256"]), (item["pdfFile"], item["pdfSha256"])]
        files += [("reference/" + image["file"], image["sha256"]) for image in item["images"]]
        for name, expected in files: assert digest(SAMPLES / name) == expected, "Saved reference differs: " + name


def main():
    parser = argparse.ArgumentParser(description=__doc__); parser.add_argument("--record", action="store_true")
    args = parser.parse_args()
    pinned = json.loads((ROOT / ".cache/acceptance-office/environment.json").read_text())
    archive = ROOT / ".cache/LibreOffice_25.8.1.1_Linux_x86-64_deb.tar.gz"
    expected_digest = "f367c434aee8352e9742e192bedcdaef194ffcbcfeaf760c6300fb4fecbba73c"
    assert pinned["sha256"] == expected_digest and digest(archive) == expected_digest, "Pinned Office archive differs"
    os.environ["PATH"] = str(ROOT / ".cache/acceptance-office/bin") + os.pathsep + os.environ["PATH"]
    current = shared.environment()
    assert current["libreOffice"] == pinned["version"] and current["libreOffice"].startswith("LibreOffice 25.8.1.1 ")
    current["officeArchiveSha256"] = expected_digest
    dependency_record = ROOT / ".cache/acceptance-office-deps/environment.json"
    if dependency_record.exists(): current["isolatedDependencies"] = json.loads(dependency_record.read_text())
    expected = json.loads((SAMPLES / "expected.json").read_text())
    baseline = None if args.record else json.loads((SAMPLES / "reference.json").read_text())
    daily = os.environ.get("DOFFI_CI_MODE") == "daily" and not args.record
    selected = [case for case in expected["cases"] if not daily or not case["file"].startswith("15-")]
    if baseline: verify_saved(baseline)
    CACHE.mkdir(parents=True, exist_ok=True); records = []
    with TemporaryDirectory(prefix="run-", dir=CACHE) as temporary:
        work = Path(temporary)
        for name in ["producer", "reopened"]: (work / name).mkdir()
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        for case in selected:
            source = SAMPLES / case["file"]; assert digest(source) == case["sha256"]
            pdf = work / (source.stem + ".pdf"); producer = work / "producer" / source.name
            run(["libreoffice", profile, "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work), str(source)])
            run(["libreoffice", profile, "--headless", "--convert-to", "xlsx:Calc MS Excel 2007 XML", "--outdir", str(producer.parent), str(source)])
            run(["libreoffice", profile, "--headless", "--convert-to", "pdf:calc_pdf_Export", "--outdir", str(work / "reopened"), str(producer)])
            page_counts = []
            for candidate in [pdf, work / "reopened" / pdf.name]:
                assert candidate.is_file(), "Calc did not produce " + str(candidate)
                info = run(["pdfinfo", str(candidate)])
                pages = int(re.search(r"^Pages:\s+(\d+)$", info, re.M)[1]); assert 1 <= pages <= 20
                page_counts.append(pages)
                text = re.sub(r"\s+", "", run(["pdftotext", "-raw", str(candidate), "-"]))
                for required in case.get("referenceText", []): assert re.sub(r"\s+", "", required) in text, (source.name, required)
            assert page_counts[0] == page_counts[1], "Resave changed print pagination: " + source.name
            run(["pdftoppm", "-r", "96", "-png", str(pdf), str(work / source.stem)])
            images = sorted(work.glob(source.stem + "-[0-9]*.png")); assert len(images) == page_counts[0]
            original, produced = features(source), features(producer)
            assert original["sheets"] == produced["sheets"], "External resave lost sheets"
            differences = {key: {"source": value, "producer": produced[key]} for key, value in original.items() if value != produced[key]}
            record = {"file": source.name, "sha256": digest(source), "pages": page_counts[0], "images": [{"file": image.name, "sha256": digest(image)} for image in images],
                      "producerFile": "producer/" + source.name, "producerSha256": digest(producer), "pdfFile": "reference/" + pdf.name, "pdfSha256": digest(pdf),
                      "sourceFeatures": original, "producerFeatures": produced, "externalDifferences": differences}
            records.append(record)
            if args.record:
                for name in ["producer", "reference"]: (SAMPLES / name).mkdir(exist_ok=True)
                for image in images: shutil.copyfile(image, SAMPLES / "reference" / image.name)
                shutil.copyfile(pdf, SAMPLES / "reference" / pdf.name); shutil.copyfile(producer, SAMPLES / "producer" / source.name)
            print(source.name, page_counts[0], "pages; external differences:", differences, flush=True)
    result = {"environment": current, "fixtures": records, "limits": ["Synthetic independent public sources; no user data.", "External feature differences are recorded, not counted as lossless success.", "PDF does not exercise comment UI, hidden sheets or cells outside selected print areas.", "Unicode text remains exact in source; emoji/variation glyph substitution is renderer and font dependent.", "Excel and WPS have not been verified."]}
    if args.record:
        (SAMPLES / "reference.json").write_text(json.dumps(result, ensure_ascii=False, indent=2) + "\n"); status = "recorded"
    else:
        saved_cases = [case for case in baseline["fixtures"] if not daily or not case["file"].startswith("15-")]
        assert len(records) == len(saved_cases)
        for actual, saved in zip(records, saved_cases):
            # Native structure must remain stable even when fallback fonts or
            # PDF tools differ. Only pixel/pagination comparisons need the
            # complete recorded rendering environment.
            for key in ["file", "sha256", "sourceFeatures", "producerFeatures", "externalDifferences"]:
                assert actual[key] == saved[key], (actual["file"], key)
            if current == baseline["environment"]:
                for key in ["pages", "images"]: assert actual[key] == saved[key], (actual["file"], key)
        status = "pixels-equal" if current == baseline["environment"] else "environment-differs-pixels-not-compared"
    (CACHE / "result.json").write_text(json.dumps({"status": status, "mode": "daily" if daily else "full", "deferredFixtures": [case["file"] for case in expected["cases"] if case not in selected], **result}, ensure_ascii=False, indent=2) + "\n")
    print(len(records), "XLSX sources opened, resaved and reopened;", status)


if __name__ == "__main__": main()
