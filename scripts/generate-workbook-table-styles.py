"""Extract all 60 normative table styles from ECMA-376 Part 1 (2016), Annex G.

Input is presetTableStyles.xml inside OfficeOpenXML-SpreadsheetMLStyles.zip.
Only standard data is used; table layout and style composition are first-party.
Ordinary XLSX dxfId values stay zero-based. This pinned preset attachment uses
one-based dxfId values within each local dxfs collection, checked explicitly.
"""
from pathlib import Path
import argparse
import hashlib
import json
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
SHA = "74c6cf73d551bb2bd33a51690d4922095260b6647b2a38fcbc0d6da6c6275d7e"
REGIONS = ["wholeTable", "firstColumnStripe", "secondColumnStripe", "firstRowStripe", "secondRowStripe", "lastColumn", "firstColumn", "headerRow", "totalRow", "firstHeaderCell", "lastHeaderCell", "firstTotalCell", "lastTotalCell"]


def check(node, attributes=(), children=()):
    if set(node.attrib) - set(attributes) or set(c.tag for c in node) - set(children):
        raise ValueError("Unmapped normative table style data: " + node.tag)


def color(node):
    check(node, ["theme", "tint"])
    theme = int(node.attrib["theme"])
    if theme not in [0, 1, 4, 5, 6, 7, 8, 9]:
        raise ValueError("Unexpected table style theme slot")
    return {"theme": theme, "tint": float(node.get("tint", "0"))}


def differential(node):
    check(node, [], ["font", "fill", "border"])
    result = {}
    for child in node:
        if child.tag == "font":
            check(child, [], ["b", "color"])
            result["font"] = {}
            bold = child.find("b")
            if bold is not None:
                check(bold)
                result["font"]["bold"] = True
            value = child.find("color")
            if value is not None:
                result["font"]["color"] = color(value)
        elif child.tag == "fill":
            check(child, [], ["patternFill"])
            pattern = child.find("patternFill")
            check(pattern, ["patternType"], ["fgColor", "bgColor"])
            if pattern.get("patternType") != "solid" or pattern.find("fgColor").attrib != pattern.find("bgColor").attrib:
                raise ValueError("A predefined fill requires another paint representation")
            result["fill"] = color(pattern.find("fgColor"))
        else:
            check(child, [], ["top", "bottom", "left", "right", "vertical", "horizontal"])
            borders = {}
            for edge in child:
                check(edge, ["style"], ["color"])
                style = edge.get("style", "none")
                if style not in ["none", "thin", "medium", "thick", "double"]:
                    raise ValueError("Unexpected predefined border style")
                borders[edge.tag] = {"style": style}
                value = edge.find("color")
                if value is not None:
                    borders[edge.tag]["color"] = color(value)
            result["borders"] = borders
    return result


def generate(source):
    raw = source.read_bytes()
    if hashlib.sha256(raw).hexdigest() != SHA:
        raise ValueError("Unexpected normative table style source digest")
    root = ET.fromstring(raw)
    names = [f"TableStyle{family}{number}" for family, count in [("Light", 21), ("Medium", 28), ("Dark", 11)] for number in range(1, count + 1)]
    definitions, indices, styles, total = [], {}, {}, 0
    compact = lambda value: json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    for name in names:
        node = root.find(name)
        check(node, [], ["dxfs", "tableStyles"])
        dxfs = node.find("dxfs")
        check(dxfs, ["count"], ["dxf"])
        # The pinned attachment says count=10 for Light14 but contains nine
        # records, all referenced exactly once. Admit only that documented
        # attachment defect; native XLSX count validation stays unchanged.
        count_defect = name == "TableStyleLight14" and dxfs.attrib["count"] == "10" and len(dxfs) == 9
        if int(dxfs.attrib["count"]) != len(dxfs) and not count_defect:
            raise ValueError("Invalid normative differential count")
        total += len(dxfs)
        local = []
        for value in dxfs:
            data = differential(value)
            key = compact(data)
            if key not in indices:
                indices[key] = len(definitions)
                definitions.append(data)
            local.append(indices[key])
        table_styles = node.find("tableStyles")
        check(table_styles, ["count", "defaultTableStyle", "defaultPivotStyle"], ["tableStyle"])
        if int(table_styles.attrib["count"]) != 1 or len(table_styles) != 1:
            raise ValueError("A preset must contain one table style")
        table = table_styles.find("tableStyle")
        check(table, ["name", "pivot", "count"], ["tableStyleElement"])
        # Two embedded names have attachment typos. The outer element and
        # Annex G.1 supply the canonical name; never relax native name checks.
        embedded_name = {"TableStyleLight14": "TableStyleLight14 ", "TableStyleDark9": "TableStyleDark9 2"}.get(name, name)
        if table.attrib["name"] != embedded_name or table.attrib["pivot"] != "0" or int(table.attrib["count"]) != len(table):
            raise ValueError("Invalid normative table style identity")
        elements, seen = [], set()
        for element in table:
            check(element, ["type", "dxfId", "size"])
            region = element.attrib["type"]
            if region not in REGIONS or region in seen:
                raise ValueError("Unexpected or duplicate table region")
            seen.add(region)
            index = int(element.attrib["dxfId"])
            if not 1 <= index <= len(local):
                raise ValueError("Preset attachment dxfId is not in its one-based local range")
            elements.append({"region": region, "size": int(element.get("size", "1")), "differential": local[index - 1]})
        if {int(e.attrib["dxfId"]) for e in table} != set(range(1, len(local) + 1)):
            raise ValueError("Preset differentials are not fully referenced")
        styles[name] = sorted(elements, key=lambda element: REGIONS.index(element["region"]))
    if total != 434:
        raise ValueError("Unexpected count of normative table differentials")
    notice = (ROOT / "packages/render/ECMA-376-TABLE-STYLES-LICENSE.txt").read_text()
    header = "/*! @license\n" + "\n".join(" * " + line for line in notice.splitlines()) + "\n */\n"
    header += '// Generated standard data; source digest and full validation are pinned in the generator.\n'
    header += 'import type { WorkbookBuiltinTableDifferential, WorkbookBuiltinTableElement } from "./table-style-types.js";\n'
    header += "export const builtinTableDifferentials: readonly WorkbookBuiltinTableDifferential[] = [\n"
    header += "\n".join("  " + compact(data) + "," for data in definitions) + "\n];\n"
    header += "export const builtinTableElements: Readonly<Record<string, readonly WorkbookBuiltinTableElement[]>> = {\n"
    header += "\n".join("  " + json.dumps(name) + ":" + compact(elements) + "," for name, elements in styles.items()) + "\n};\n"
    return header, len(definitions)


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    target = ROOT / "packages/render/src/workbook/table-style-data.ts"
    generated, count = generate(args.source)
    if args.check:
        if generated != target.read_text():
            raise ValueError("Generated table style data differs")
    else:
        target.write_text(generated)
    print("Verified" if args.check else "Generated", 60, "table styles,", count, "unique differentials from 434 source records.")
