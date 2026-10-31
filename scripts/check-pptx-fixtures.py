"""Independent fixture integrity and relationship/semantic checks (no SDK)."""
from io import BytesIO
from pathlib import Path
from tempfile import TemporaryDirectory
from urllib.parse import unquote
from zipfile import ZipFile
import hashlib
import importlib.util
import json
import posixpath
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
spec = importlib.util.spec_from_file_location("pptx_generator", ROOT / "scripts/generate-pptx-fixtures.py")
g = importlib.util.module_from_spec(spec)
spec.loader.exec_module(g)
samples = ROOT / "examples/pptx"
expected = json.loads((samples / "expected.json").read_text())


def require(value, message):
    if not value:
        raise AssertionError(message)


def files(path):
    with ZipFile(path) as archive:
        require(len(archive.namelist()) == len(set(archive.namelist())), "Duplicate ZIP entry")
        return {name: archive.read(name) for name in archive.namelist()}


def audit(parts):
    roots = {name: ET.fromstring(data) for name, data in parts.items()
             if name.endswith(".xml") or name.endswith(".rels")}
    content = roots["[Content_Types].xml"]
    types = {unquote(item.get("PartName", "").lstrip("/")): item.get("ContentType")
             for item in content if item.tag.endswith("}Override")}
    defaults = {item.get("Extension"): item.get("ContentType")
                for item in content if item.tag.endswith("}Default")}
    for name, root in roots.items():
        require(types.get(name) or defaults.get(name.rsplit(".", 1)[-1]), "Missing content type: " + name)
        if name.endswith(".rels"):
            source = "" if name == "_rels/.rels" else posixpath.dirname(name)[:-5] + posixpath.basename(name)[:-5]
            require(not source or source in parts, "Missing relationship owner")
            ids = set()
            for relationship in root:
                rid = relationship.get("Id")
                require(rid and rid not in ids, "Invalid relationship identity")
                ids.add(rid)
                if relationship.get("TargetMode") == "External":
                    require(relationship.get("Type").endswith("/hyperlink"), "Unexpected external resource")
                    continue
                target = unquote(relationship.get("Target"))
                target = target.lstrip("/") if target.startswith("/") else posixpath.normpath(posixpath.join(posixpath.dirname(source), target))
                require(target in parts, "Missing relationship target: " + target)
        if root.tag.endswith(("}sld", "}sldMaster", "}sldLayout", "}notes", "}notesMaster")):
            ids = [node.get("id") for node in root.iter() if node.tag.endswith("}cNvPr")]
            require(len(ids) == len(set(ids)), "Duplicate native object ID")
    return roots


with TemporaryDirectory() as temporary:
    generated = Path(temporary)
    g.generate(generated)
    require((generated / "expected.json").read_bytes() == (samples / "expected.json").read_bytes(), "Expectations are not reproducible")
    for case in expected["cases"] + expected["negativeCases"]:
        path = samples / case["file"]
        data = path.read_bytes()
        require(hashlib.sha256(data).hexdigest() == case["sha256"], "Fixture digest differs")
        require((generated / case["file"]).read_bytes() == data, "Generator bytes differ: " + case["file"])

for case in expected["cases"]:
    parts = files(samples / case["file"])
    roots = audit(parts)
    source_text = "\n".join("".join(node.itertext()) for node in roots.values())
    for required in case["text"] + case.get("notes", []) + case.get("commentText", []):
        require(required in source_text, "Source text missing: " + required)
    for name, data in parts.items():
        if name.endswith(".xlsx"):
            embedded = audit(files(BytesIO(data)))
            require(not any(node.tag.endswith("}f") for root in embedded.values() for node in root.iter()), "Chart data has a formula")
    if "modernComments" in case:
        author_ids = {node.get("id") for root in roots.values() for node in root.iter() if node.tag == "{" + g.MODERN + "}author"}
        comments = [node for root in roots.values() for node in root.iter() if node.tag == "{" + g.MODERN + "}cm"]
        require(len(comments) == case["modernComments"], "Modern thread count differs")
        for node in comments + [node for root in roots.values() for node in root.iter() if node.tag == "{" + g.MODERN + "}reply"]:
            require(node.get("authorId") in author_ids, "Unknown modern author")
    if "textAnchor" in case:
        anchor = case["textAnchor"]
        value = case["text"][-1].encode("utf-16-le")
        require(value[anchor["cp"] * 2:(anchor["cp"] + anchor["length"]) * 2].decode("utf-16-le") == anchor["quote"], "UTF-16 anchor differs")

negatives = {case["file"]: files(samples / case["file"]) for case in expected["negativeCases"] if "truncated" not in case["file"]}
for name in ["missing-slide", "duplicate-native-id", "external-image"]:
    try:
        audit(negatives["invalid/" + name + ".pptx"])
    except AssertionError:
        pass
    else:
        raise AssertionError("Negative fixture is ineffective: " + name)
formula = files(BytesIO(negatives["invalid/chart-formula.pptx"]["ppt/embeddings/data1.xlsx"]))
require(b"<f>1+1</f>" in formula["xl/worksheets/sheet1.xml"], "Formula negative fixture is ineffective")
require(b"<!DOCTYPE" in negatives["invalid/doctype.pptx"]["ppt/slides/slide1.xml"], "DTD negative fixture is ineffective")
require(b"<p:timing>" in negatives["invalid/animation.pptx"]["ppt/slides/slide1.xml"], "Animation negative fixture is ineffective")
require(b"unsupported-smartart" in negatives["invalid/unsupported-object.pptx"]["ppt/slides/slide1.xml"], "Graphic negative fixture is ineffective")
require(b'TargetMode="External"' in negatives["invalid/external-chart-data.pptx"]["ppt/charts/_rels/chart1.xml.rels"], "External data negative fixture is ineffective")
require(b'authorId="{99999999-' in negatives["invalid/unknown-comment-author.pptx"]["ppt/comments/modernComment1.xml"], "Author negative fixture is ineffective")
try:
    files(samples / "invalid/truncated-zip.pptx")
except Exception as error:
    from zipfile import BadZipFile
    require(isinstance(error, BadZipFile), "Unexpected malformed ZIP result")
else:
    raise AssertionError("ZIP negative fixture is ineffective")
print("Verified 16 deterministic independent PPTX sources, 10 negative cases, native relationships, chart constants and UTF-16 comment anchor.")
