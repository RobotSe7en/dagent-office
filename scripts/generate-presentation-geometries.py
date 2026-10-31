"""Extract selected normative geometry data; no Office/editor implementation code.

Obtain presetShapeDefinitions.xml from the official ECMA-376 Part 1 (2016)
DrawingML geometries annex. The pinned input digest prevents silent data changes.
"""
from pathlib import Path
import argparse
import hashlib
import json
import xml.etree.ElementTree as ET

ROOT = Path(__file__).resolve().parent.parent
SHA = "2f7c868d857c1e3c4b5a6068759fe0e07d77ad58377a6618d1b02ba3507b6939"
A = "{http://schemas.openxmlformats.org/drawingml/2006/main}"
NAMES = ["rect", "roundRect", "ellipse", "triangle", "diamond", "parallelogram", "trapezoid", "pentagon", "hexagon", "rightArrow", "leftArrow", "downArrow", "leftRightArrow", "chevron", "line", "straightConnector1", "bentConnector2", "bentConnector3", "bentConnector4", "bentConnector5"]


def generate(source):
    raw = source.read_bytes()
    if hashlib.sha256(raw).hexdigest() != SHA:
        raise ValueError("Unexpected normative geometry source digest")
    root = ET.fromstring(raw)
    data = {}
    for name in NAMES:
        node = root.find(name)
        guides = lambda tag: [[g.attrib["name"], g.attrib["fmla"]] for g in node.findall(A + tag + "/" + A + "gd")]
        rectangle = node.find(A + "rect")
        paths = []
        for path in node.findall(A + "pathLst/" + A + "path"):
            if set(path.attrib) - {"fill", "stroke", "extrusionOk"}:
                raise ValueError("A selected path requires unsupported coordinate-space mapping")
            commands = []
            for command in path:
                kind = command.tag.removeprefix(A)
                if kind in ["moveTo", "lnTo"]:
                    point = command.find(A + "pt")
                    commands.append(["M" if kind == "moveTo" else "L", point.attrib["x"], point.attrib["y"]])
                elif kind == "arcTo":
                    commands.append(["A", *[command.attrib[key] for key in ["wR", "hR", "stAng", "swAng"]]])
                elif kind == "close":
                    commands.append(["Z"])
                else:
                    raise ValueError("Unexpected selected path command: " + kind)
            paths.append({"fill": path.get("fill", "norm") != "none", "stroke": path.get("stroke", "true") != "false", "commands": commands})
        data[name] = {"adjustments": guides("avLst"), "guides": guides("gdLst"),
                      "text": [rectangle.attrib[key] for key in ["l", "t", "r", "b"]] if rectangle is not None else None,
                      "connections": [[c.attrib["ang"], c.find(A + "pos").attrib["x"], c.find(A + "pos").attrib["y"]] for c in node.findall(A + "cxnLst/" + A + "cxn")],
                      "paths": paths}
    notice = (ROOT / "packages/core/ECMA-376-GEOMETRY-LICENSE.txt").read_text()
    header = "/*! @license\n" + "\n".join(" * " + line for line in notice.splitlines()) + "\n */\n"
    types = """// Generated standard data, not source document content. See the generator and license.
export interface PresetDefinition {
  readonly adjustments: readonly (readonly [string, string])[];
  readonly guides: readonly (readonly [string, string])[];
  readonly text: readonly [string, string, string, string] | null;
  readonly connections: readonly (readonly [string, string, string])[];
  readonly paths: readonly { readonly fill: boolean; readonly stroke: boolean; readonly commands: readonly (readonly string[])[] }[];
}
"""
    return types + header + "export const presetDefinitions: Readonly<Record<string, PresetDefinition>> = " + json.dumps(data, ensure_ascii=False, indent=2) + ";\n"


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("source", type=Path)
    parser.add_argument("--check", action="store_true")
    args = parser.parse_args()
    target = ROOT / "packages/core/src/presentation-geometry-data.ts"
    generated = generate(args.source)
    if args.check:
        if generated != target.read_text():
            raise ValueError("Generated preset data differs")
    else:
        target.write_text(generated)
    print("Verified" if args.check else "Generated", len(NAMES), "selected normative geometry definitions.")
