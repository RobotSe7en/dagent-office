"""Independent official ECMA Transitional XSD plus selected MS-PPTX comment checks.

Schemas are unmodified standard electronic inserts, never SDK implementation code.
The modern profile checks the selected anchor/author/thread/status vocabulary and
validates every rich body with the standard DrawingML CT_TextBody type.
"""
from copy import deepcopy
from hashlib import sha256
from io import BytesIO
import json
from pathlib import Path
import re
import sys
from tempfile import TemporaryDirectory
from zipfile import ZipFile
from lxml import etree as E

ROOT = Path(__file__).resolve().parent.parent
ARCHIVE = ROOT / "scripts/fixtures/ooxml/OfficeOpenXML-XMLSchema-Transitional.zip"
EXPECTED = "d34187520749998af306faf1b730e568b0ca6d88ad24638a407c0a9bb4ca04fc"
P = "http://schemas.openxmlformats.org/presentationml/2006/main"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
S = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
M = "http://schemas.microsoft.com/office/powerpoint/2018/8/main"
PC = "http://schemas.microsoft.com/office/powerpoint/2013/main/command"
AC = "http://schemas.microsoft.com/office/drawing/2013/main/command"
GUID = re.compile(r"\{[\dA-Fa-f]{8}-(?:[\dA-Fa-f]{4}-){3}[\dA-Fa-f]{12}\}")


def check_modern(root, body_schema):
    tag = E.QName(root).localname
    if tag == "authorLst":
        for author in root:
            assert author.tag == "{" + M + "}author"
            assert GUID.fullmatch(author.get("id", ""))
            assert all(name in author.attrib for name in ["name", "userId", "providerId"])
        return
    assert tag == "cmLst"
    identities = set()
    for comment in root:
        assert comment.tag == "{" + M + "}cm"
        for message in [comment] + comment.findall("{" + M + "}replyLst/{" + M + "}reply"):
            assert GUID.fullmatch(message.get("id", "")) and message.get("id") not in identities
            identities.add(message.get("id"))
            assert GUID.fullmatch(message.get("authorId", ""))
            assert message.get("status", "active") in ["active", "resolved", "closed"]
            assert re.fullmatch(r"\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(?:\.\d+)?(?:Z|[+-]\d\d:\d\d)?", message.get("created", ""))
            for body in message.findall("{" + M + "}txBody"):
                clone = deepcopy(body)
                clone.tag = "{urn:dagent-office:validation}body"
                body_schema.assertValid(clone)
        children = [E.QName(child) for child in comment]
        anchor = [child for child in children if child.namespace in [PC, AC] or child.localname == "unknownAnchor"]
        assert anchor
        kinds = {(child.namespace, child.localname) for child in anchor}
        assert len(kinds) == 1
        namespace, name = next(iter(kinds))
        assert (namespace, name) in [(PC, "sldMkLst"), (AC, "deMkLst"), (AC, "txMkLst"), (M, "unknownAnchor")]
        if name in ["sldMkLst", "unknownAnchor"]:
            assert len(anchor) == 1
        positions = {name: index for index, name in enumerate(["anchor", "pos", "replyLst", "txBody", "extLst"])}
        order = [positions["anchor" if child in anchor else child.localname] for child in children]
        assert order == sorted(order)
        for moniker in comment.iter():
            q = E.QName(moniker)
            if q.namespace == PC and q.localname == "sldMk":
                assert 256 <= int(moniker.get("sldId")) <= 2147483647
            if q.namespace == AC and q.localname == "txMk":
                assert int(moniker.get("cp")) >= 0 and int(moniker.get("len", "0")) >= 0
            if q.namespace == AC and q.localname == "context":
                assert 0 <= int(moniker.get("hash")) <= 4294967295


def check(directory):
    assert sha256(ARCHIVE.read_bytes()).hexdigest() == EXPECTED
    cases = []
    with TemporaryDirectory(prefix="dagent-native-schemas-") as temporary:
        path = Path(temporary)
        with ZipFile(ARCHIVE) as archive:
            for name in archive.namelist():
                assert "/" not in name and "\\" not in name
                (path / name).write_bytes(archive.read(name))
        parser = E.XMLParser(resolve_entities=False, no_network=True)
        schemas = {namespace: E.XMLSchema(E.parse(str(path / (name + ".xsd")), parser))
                   for namespace, name in [(P, "pml"), (A, "dml-main"), (C, "dml-chart"), (S, "sml")]}
        wrapper = path / "body.xsd"
        wrapper.write_text('<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:a="' + A +
                           '" targetNamespace="urn:dagent-office:validation" elementFormDefault="qualified">'
                           '<xsd:import namespace="' + A + '" schemaLocation="dml-main.xsd"/>'
                           '<xsd:element name="body" type="a:CT_TextBody"/></xsd:schema>')
        body_schema = E.XMLSchema(E.parse(str(wrapper), parser))
        for file in sorted(directory.glob("*.pptx")):
            counts = {"standardParts": 0, "modernParts": 0, "embeddedParts": 0}
            def parts(archive, embedded=False):
                for name in archive.namelist():
                    if name.endswith(".xlsx"):
                        with ZipFile(BytesIO(archive.read(name))) as workbook:
                            parts(workbook, True)
                    if not name.endswith(".xml"):
                        continue
                    root = E.fromstring(archive.read(name), parser)
                    namespace = E.QName(root).namespace
                    if namespace in schemas:
                        try:
                            schemas[namespace].assertValid(root)
                        except E.DocumentInvalid as error:
                            raise AssertionError(str(file) + " / " + name + ": " + str(error)) from error
                        counts["embeddedParts" if embedded else "standardParts"] += 1
                    elif namespace == M:
                        check_modern(root, body_schema)
                        counts["modernParts"] += 1
            with ZipFile(file) as archive:
                parts(archive)
            cases.append({"file": file.name, **counts})
    result = {"schemasSha256": EXPECTED, "lxml": E.LXML_VERSION, "cases": cases}
    (directory.parent / "schema-report.json").write_text(json.dumps(result, indent=2) + "\n")
    print("Official Transitional XSD and selected modern comment checks:", len(cases), "PPTX files passed.")


if __name__ == "__main__":
    check(Path(sys.argv[1]))
