import type { WorkbookDrawingAnchor, WorkbookDrawingMarker } from "@dagent-office/core";
import { check, fail, text } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, boolean, one, value } from "./values.js";

export const XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
const minCoordinate = -27273042329600, maxCoordinate = 27273042316900;
function coordinate(source: string | undefined, node: XmlNode, isExtent = false): number {
  const units = source === undefined || isExtent ? null : /^(-?\d+(?:\.\d+)?)(mm|cm|in|pt|pc|pi)$/.exec(source.trim());
  if (units) {
    const factors: Record<string, number> = { mm: 72 / 25.4, cm: 72 / 2.54, in: 72, pt: 1, pc: 12, pi: 12 };
    const points = Number(units[1]) * factors[units[2]!]!;
    if (!Number.isFinite(points) || points < minCoordinate / 12700 || points > maxCoordinate / 12700) fail(node, "A measured drawing coordinate exceeds its supported bounds.");
    return points;
  }
  if (source === undefined || !/^[+-]?\d+$/.test(source.trim())) fail(node, "A drawing coordinate must be an integer EMU value.");
  const number = Number(source);
  if (!Number.isSafeInteger(number) || number < (isExtent ? 0 : minCoordinate) || number > maxCoordinate) fail(node, "A drawing coordinate exceeds its supported native bounds.");
  return number / 12700;
}
export function drawingMarker(node: XmlNode): WorkbookDrawingMarker {
  attributes(node, []); check(node, ["row", "col", "rowOff", "colOff"], XDR);
  const row = one(node, "row", XDR), column = one(node, "col", XDR), y = one(node, "rowOff", XDR), x = one(node, "colOff", XDR);
  if (!row || !column || !x || !y) fail(node, "A drawing marker is missing its coordinate or offset.");
  for (const field of [row, column, x, y]) {
    attributes(field, []); if (field.content.some(item => typeof item !== "string")) fail(field, "A drawing marker value must contain only text.");
  }
  const index = (field: XmlNode, maximum: number): number => {
    const raw = text(field).trim();
    if (!/^[+]?\d+$/.test(raw) || !Number.isSafeInteger(Number(raw)) || Number(raw) > maximum) fail(field, "A drawing cell marker exceeds worksheet bounds.");
    return Number(raw);
  };
  return { row: index(row, 1048575), column: index(column, 16383), offsetXPt: coordinate(text(x), x), offsetYPt: coordinate(text(y), y) };
}
function extent(node: XmlNode | undefined, parent: XmlNode): { widthPt: number; heightPt: number } {
  if (!node) fail(parent, "A drawing anchor is missing its extent.");
  attributes(node, ["cx", "cy"]); check(node, [], XDR);
  return { widthPt: coordinate(value(node, "cx"), node, true), heightPt: coordinate(value(node, "cy"), node, true) };
}
export interface NativeDrawingAnchor {
  readonly anchor: WorkbookDrawingAnchor;
  readonly locked: boolean;
  readonly printable: boolean;
  /** A feature-specific adapter must consume this object after its anchor. */
  readonly object: XmlNode;
}
/** Geometry and row/column response share one source representation. Native
 * object xfrm offsets/extents are display caches for sheet-anchored objects. */
export function readDrawingAnchor(node: XmlNode): NativeDrawingAnchor {
  if (node.uri !== XDR || !["absoluteAnchor", "oneCellAnchor", "twoCellAnchor"].includes(node.name)) fail(node, "A worksheet drawing anchor has an invalid root.");
  attributes(node, node.name === "twoCellAnchor" ? ["editAs"] : []);
  const geometry = node.name === "absoluteAnchor" ? ["pos", "ext"] : node.name === "oneCellAnchor" ? ["from", "ext"] : ["from", "to"];
  check(node, [...geometry, "sp", "pic", "graphicFrame", "grpSp", "cxnSp", "contentPart", "clientData"], XDR);
  const objects = node.content.filter((item): item is XmlNode => typeof item !== "string" && !geometry.includes(item.name) && item.name !== "clientData");
  if (objects.length !== 1) fail(node, "A worksheet anchor must contain exactly one drawing object.");
  const client = one(node, "clientData", XDR); if (!client) fail(node, "A worksheet drawing anchor has no client data.");
  attributes(client, ["fLocksWithSheet", "fPrintsWithSheet"]); check(client, [], XDR);
  let anchor: WorkbookDrawingAnchor;
  if (node.name === "absoluteAnchor") {
    const position = one(node, "pos", XDR); if (!position) fail(node, "An absolute drawing anchor has no position.");
    attributes(position, ["x", "y"]); check(position, [], XDR);
    anchor = { kind: "absolute", xPt: coordinate(value(position, "x"), position), yPt: coordinate(value(position, "y"), position), ...extent(one(node, "ext", XDR), node) };
  } else {
    const from = one(node, "from", XDR); if (!from) fail(node, "A worksheet drawing anchor has no starting marker.");
    if (node.name === "oneCellAnchor") anchor = { kind: "one-cell", from: drawingMarker(from), ...extent(one(node, "ext", XDR), node) };
    else {
      const to = one(node, "to", XDR), mode = value(node, "editAs") ?? "twoCell";
      if (!to || !["twoCell", "oneCell", "absolute"].includes(mode)) fail(node, "A two-cell anchor has no endpoint or an invalid editing mode.");
      anchor = { kind: "two-cell", from: drawingMarker(from), to: drawingMarker(to), moveWithCells: mode !== "absolute", sizeWithCells: mode === "twoCell" };
    }
  }
  return { anchor, object: objects[0]!, locked: boolean(client, "fLocksWithSheet", true), printable: boolean(client, "fPrintsWithSheet", true) };
}
/** ISO 29500 commentPr uses the same markers and independent response flags. */
export function readNoteObjectAnchor(node: XmlNode): WorkbookDrawingAnchor {
  attributes(node, ["moveWithCells", "sizeWithCells"]); check(node, ["from", "to"], XDR);
  const from = one(node, "from", XDR), to = one(node, "to", XDR);
  if (!from || !to) fail(node, "A note object anchor is missing an endpoint.");
  return { kind: "two-cell", from: drawingMarker(from), to: drawingMarker(to), moveWithCells: boolean(node, "moveWithCells"), sizeWithCells: boolean(node, "sizeWithCells") };
}
/** VML Anchor offsets are pixels, not the binary XLS fractional cell units.
 * Convert source units at the format boundary, preserving overflow offsets. */
export function readVmlAnchor(node: XmlNode, moveWithCells: boolean, sizeWithCells: boolean): WorkbookDrawingAnchor {
  attributes(node, []);
  if (node.content.some(item => typeof item !== "string")) fail(node, "A VML note anchor must contain only text.");
  const tokens = text(node).split(",").map(item => item.trim());
  if (tokens.length !== 8 || tokens.some(item => !/^[+-]?\d+$/.test(item) || !Number.isSafeInteger(Number(item)))) fail(node, "A VML note anchor requires eight integer components.");
  const values = tokens.map(Number);
  const marker = (offset: number): WorkbookDrawingMarker => {
    const column = values[offset]!, row = values[offset + 2]!, x = values[offset + 1]!, y = values[offset + 3]!;
    // VML pixel offsets use a bounded signed integer representation. They may span more than
    // one cell and must not be truncated to the current viewport's row size.
    if (row < 0 || row > 1048575 || column < 0 || column > 16383 || [x, y].some(v => v < -2147483648 || v > 2147483647)) fail(node, "A VML note anchor exceeds native coordinate bounds.");
    return { row, column, offsetXPt: x * 0.75, offsetYPt: y * 0.75 };
  };
  return { kind: "two-cell", from: marker(0), to: marker(4), moveWithCells, sizeWithCells };
}
