"""Reproduce and independently audit source XLSX fixtures without the SDK."""
from pathlib import Path
from tempfile import TemporaryDirectory
from zipfile import ZipFile, BadZipFile
import hashlib
import importlib.util
import json

ROOT = Path(__file__).resolve().parent.parent
SAMPLES = ROOT / "examples/xlsx"


def module(name):
    spec = importlib.util.spec_from_file_location(name, ROOT / "scripts" / (name + ".py"))
    result = importlib.util.module_from_spec(spec); spec.loader.exec_module(result)
    return result


g = module("generate-xlsx-fixtures")
check = module("check-xlsx-output")
expected = json.loads((SAMPLES / "expected.json").read_text())
S, TC, D = check.S, check.TC, check.D

with TemporaryDirectory(prefix="dagent-xlsx-fixtures-") as directory:
    generated = Path(directory); g.generate(generated)
    assert (generated / "expected.json").read_bytes() == (SAMPLES / "expected.json").read_bytes(), "Expectations are not reproducible"
    for case in expected["cases"] + expected["negativeCases"]:
        data = (SAMPLES / case["file"]).read_bytes()
        assert hashlib.sha256(data).hexdigest() == case["sha256"], case["file"]
        assert data == (generated / case["file"]).read_bytes(), "Source bytes differ: " + case["file"]

with check.schemas() as standards:
    for case in expected["cases"]:
        roots, relations = check.audit(SAMPLES / case["file"], standards)
        wb_path, wb = next((name, root) for name, root in roots.items() if root.tag == "{" + S + "}workbook")
        sheets = wb.find("{" + S + "}sheets")
        assert [s.get("name") for s in sheets] == case["sheets"], case["file"]
        cells, formula_count = {}, 0
        strings = next((root for root in roots.values() if root.tag == "{" + S + "}sst"), [])
        for sheet in sheets:
            path = relations[wb_path][sheet.get("{" + check.R + "}id")][1]
            for cell in roots[path].findall("{" + S + "}sheetData/{" + S + "}row/{" + S + "}c"):
                formula_count += cell.find("{" + S + "}f") is not None
                kind = cell.get("t", "n"); value = cell.findtext("{" + S + "}v")
                if kind == "inlineStr": value = "".join(cell.find("{" + S + "}is").itertext())
                elif kind == "s": value = "".join(strings[int(value)].itertext())
                elif kind == "b": value = value == "1"
                elif kind == "n" and value is not None: value = float(value)
                cells[sheet.get("name") + "!" + cell.get("r")] = value
        assert len(cells) == case["cellCount"], (case["file"], len(cells), case["cellCount"])
        assert formula_count == case["formulaCount"], (case["file"], formula_count, case["formulaCount"])
        for address, value in case.get("values", {}).items(): assert cells[address] == value, (case["file"], address)
        source_text = "\n".join("".join(root.itertext()) for root in roots.values())
        for value in case["text"]: assert value in source_text, (case["file"], value)
        counts = {"conditionalFormats": (S, "cfRule"), "validations": (S, "dataValidation"), "classicNotes": (S, "comment"), "autoFilters": (S, "autoFilter"), "images": (D, "pic"), "shapes": (D, "sp")}
        for name, (namespace, tag) in counts.items():
            if name in case: assert sum(len(list(root.iter("{" + namespace + "}" + tag))) for root in roots.values()) == case[name], (case["file"], name)
        if "threads" in case:
            messages = [item for root in roots.values() for item in root.iter("{" + TC + "}threadedComment")]
            assert sum(item.get("parentId") is None for item in messages) == case["threads"]
            assert sum(item.get("parentId") is not None for item in messages) == case["replies"]
            assert sum(item.get("done") == "1" for item in messages) == case["resolvedThreads"]
        if "dateSystem" in case:
            assert ("1904" if wb.find("{" + S + "}workbookPr").get("date1904") == "1" else "1900") == case["dateSystem"]
        print(case["file"], len(cells), "cells;", formula_count, "formulas; native structure passed")

for case in expected["negativeCases"]:
    path = SAMPLES / case["file"]
    if path.stem == "truncated-zip":
        try: ZipFile(path)
        except BadZipFile: pass
        else: raise AssertionError("Truncated ZIP remained readable")
    else:
        with ZipFile(path) as archive:
            if "missingPart" in case: assert case["missingPart"] not in archive.namelist()
            else: assert case["marker"].encode() in archive.read(case["part"])

print(f'Independent XLSX fixtures: {len(expected["cases"])} valid sources and {len(expected["negativeCases"])} effective negative mutations passed.')
