"""Record independent PPTX references, or verify them without changing baselines.

Uses a private temporary LibreOffice profile and process group. Different renderer
environments still run open/resave/text/page checks, but never pass pixel equality.
"""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile
import argparse
import hashlib
import json
import os
import re
import shutil
import signal
import subprocess
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "examples/pptx"
CACHE = ROOT / ".cache/pptx-reference"


def digest(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(command, timeout=55):
    process = subprocess.Popen(command, cwd=ROOT, stdout=subprocess.PIPE,
                               stderr=subprocess.PIPE, text=True, start_new_session=True)
    try:
        stdout, stderr = process.communicate(timeout=timeout)
        if process.returncode:
            raise RuntimeError(f"{command[0]} failed: {stderr or stdout}")
        return stdout.strip()
    finally:
        # soffice can leave a child behind, including when its launcher exits.
        try:
            os.killpg(process.pid, signal.SIGTERM)
        except ProcessLookupError:
            pass
        if process.poll() is None:
            try:
                process.communicate(timeout=3)
            except subprocess.TimeoutExpired:
                os.killpg(process.pid, signal.SIGKILL)
                process.communicate()


def environment():
    fonts = []
    for family in ["Noto Sans CJK SC", "Liberation Sans"]:
        actual = run(["fc-match", "-f", "%{family}", family])
        if family not in actual:
            raise RuntimeError("Required reference font unavailable: " + family)
    # Record all installed fallback fonts too: they affect CJK/emoji substitution.
    for value in sorted(set(run(["fc-list", "-f", "%{file}\n"]).splitlines())):
        path = Path(value)
        fonts.append({"file": path.name, "sha256": digest(path)})
    return {"libreOffice": run(["libreoffice", "--version"]),
            "poppler": run(["pdfinfo", "-v"]) or subprocess.run(
                ["pdfinfo", "-v"], capture_output=True, text=True, check=True).stderr.strip(),
            "fontconfig": run(["fc-match", "-V"]) or subprocess.run(
                ["fc-match", "-V"], capture_output=True, text=True, check=True).stderr.strip(),
            "dpi": 96, "fonts": sorted(fonts, key=lambda item: (item["file"], item["sha256"]))}


def features(path):
    counts = {name: 0 for name in ["slides", "charts", "notes", "classicComments", "modernComments", "replies"]}
    modern = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
    p = "http://schemas.openxmlformats.org/presentationml/2006/main"
    with ZipFile(path) as archive:
        for name in archive.namelist():
            if not name.endswith(".xml"):
                continue
            root = ET.fromstring(archive.read(name))
            if root.tag.endswith("}sld"):
                counts["slides"] += 1
            elif root.tag.endswith("}chartSpace"):
                counts["charts"] += 1
            elif root.tag.endswith("}notes"):
                counts["notes"] += 1
            counts["classicComments"] += len(root.findall("./{" + p + "}cm"))
            counts["modernComments"] += len(root.findall("./{" + modern + "}cm"))
            counts["replies"] += len(root.findall(".//{" + modern + "}reply"))
    return counts


def verify_saved(record):
    expected = json.loads((SAMPLES / "expected.json").read_text())["cases"]
    if [item["file"] for item in record["fixtures"]] != [item["file"] for item in expected]:
        raise RuntimeError("Reference fixture identities or count differ from the independent corpus.")
    for item, case in zip(record["fixtures"], expected):
        if item["sha256"] != case["sha256"] or item["pages"] != case["slides"]:
            raise RuntimeError("Reference source or page count differs from independent expectations.")
        images = item["images"]
        if len(images) != case["slides"] or len({image["file"] for image in images}) != len(images):
            raise RuntimeError("Reference images are missing or duplicated.")
    for item in record["fixtures"]:
        for name, expected in [(item["file"], item["sha256"]),
                               (item["producerFile"], item["producerSha256"]),
                               (item["pdfFile"], item["pdfSha256"])]:
            if digest(SAMPLES / name) != expected:
                raise RuntimeError("Saved fixture differs: " + name)
        for image in item["images"]:
            if digest(SAMPLES / "reference" / image["file"]) != image["sha256"]:
                raise RuntimeError("Saved reference differs: " + image["file"])


def verify_features(source, producer):
    # Different Office versions may preserve modern comments or drop them; the
    # result records either outcome. Loss of supported base content always fails.
    for key in ["slides", "charts", "notes", "classicComments"]:
        if source[key] != producer[key]:
            raise RuntimeError("LibreOffice resave lost supported base content: " + key)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--record", action="store_true")
    args = parser.parse_args()
    expected = json.loads((SAMPLES / "expected.json").read_text())
    baseline = None
    if not args.record:
        baseline = json.loads((SAMPLES / "reference.json").read_text())
        verify_saved(baseline)
    CACHE.mkdir(parents=True, exist_ok=True)
    current = environment()
    records = []
    # A clean output directory prevents stale PDF/resave files from passing.
    with TemporaryDirectory(prefix="run-", dir=CACHE) as temporary:
        work = Path(temporary)
        producer = work / "producer"
        producer.mkdir()
        sources = [SAMPLES / item["file"] for item in expected["cases"]]
        profile = "-env:UserInstallation=" + (work / "profile").as_uri()
        run(["libreoffice", profile, "--headless", "--convert-to",
             'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}',
             "--outdir", str(work), *map(str, sources)])
        run(["libreoffice", profile, "--headless", "--convert-to", "pptx:Impress MS PowerPoint 2007 XML",
             "--outdir", str(producer), *map(str, sources)])
        # Reopen the resaved package too, with hidden pages explicitly enabled.
        reopened = work / "reopened"
        reopened.mkdir()
        run(["libreoffice", profile, "--headless", "--convert-to",
             'pdf:impress_pdf_Export:{"ExportHiddenSlides":{"type":"boolean","value":"true"}}',
             "--outdir", str(reopened), *map(lambda source: str(producer / source.name), sources)])
        for case, source in zip(expected["cases"], sources):
            if digest(source) != case["sha256"]:
                raise RuntimeError("Source differs from independent expectations: " + source.name)
            pdf = work / (source.stem + ".pdf")
            for candidate in [pdf, reopened / pdf.name]:
                if not candidate.is_file():
                    raise RuntimeError("LibreOffice did not produce " + str(candidate))
                info = run(["pdfinfo", str(candidate)])
                pages = int(re.search(r"^Pages:\s+(\d+)$", info, re.M).group(1))
                if pages != case["slides"]:
                    raise RuntimeError("Reference page count differs: " + source.name)
                text = re.sub(r"\s+", "", run(["pdftotext", "-raw", str(candidate), "-"]))
                for required in case.get("referenceText", case["text"]):
                    if re.sub(r"\s+", "", required) not in text:
                        raise RuntimeError(f"Reference text absent in {source.name}: {required}")
            run(["pdftoppm", "-r", "96", "-png", str(pdf), str(work / source.stem)])
            images = sorted(work.glob(source.stem + "-[0-9]*.png"))
            if len(images) != case["slides"]:
                raise RuntimeError("Reference image count differs: " + source.name)
            source_features = features(source)
            producer_features = features(producer / source.name)
            verify_features(source_features, producer_features)
            records.append({"file": source.name, "sha256": digest(source), "pages": pages,
                            "images": [{"file": image.name, "sha256": digest(image)} for image in images],
                            "producerFile": "producer/" + source.name,
                            "producerSha256": digest(producer / source.name),
                            "pdfFile": "reference/" + pdf.name, "pdfSha256": digest(pdf),
                            "sourceFeatures": source_features, "producerFeatures": producer_features,
                            "referenceLimits": case.get("referenceLimits", [])})
            if args.record:
                (SAMPLES / "reference").mkdir(exist_ok=True)
                (SAMPLES / "producer").mkdir(exist_ok=True)
                for image in images:
                    shutil.copyfile(image, SAMPLES / "reference" / image.name)
                shutil.copyfile(producer / source.name, SAMPLES / "producer" / source.name)
                shutil.copyfile(pdf, SAMPLES / "reference" / pdf.name)
    record = {"environment": current, "fixtures": records,
              "limits": ["Synthetic public fixtures; no user documents.",
                         "LibreOffice resave may drop modern comments and replies; per-file counts record this loss.",
                         "PDF excludes comment UI and notes; these are checked in source XML.",
                         "U+20000/variation glyph availability requires separate font diagnostics.",
                         "PowerPoint and WPS application acceptance has not been performed."]}
    if args.record:
        (SAMPLES / "reference.json").write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n")
        status = "recorded"
    elif current != baseline["environment"]:
        status = "environment-differs-pixels-not-compared"
    else:
        # Resave ZIP bytes include producer timestamps; compare actual semantic
        # counts and PDF pixels, not nondeterministic producer ZIP digests.
        for actual, saved in zip(records, baseline["fixtures"]):
            for key in ["file", "sha256", "pages", "images", "sourceFeatures", "producerFeatures", "referenceLimits"]:
                if actual[key] != saved[key]:
                    raise RuntimeError(f"Reference differs: {actual['file']} ({key})")
        status = "pixels-equal"
    (CACHE / "result.json").write_text(json.dumps({"status": status, **record}, ensure_ascii=False, indent=2) + "\n")
    print(f"PPTX references: {len(records)} opened, resaved and reopened; {status}.")


if __name__ == "__main__":
    main()
