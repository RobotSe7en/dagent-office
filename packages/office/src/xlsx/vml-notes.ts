import { DocumentError } from "@doffi/core";
import type { WorkbookCoordinate, WorkbookDrawingAnchor } from "@doffi/core";
import type { OpcPackage } from "../ooxml/package.js";
import { attr, check, elements, fail, R, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, one, value } from "./values.js";
import { readVmlAnchor } from "./drawing-anchor.js";

export const VML = "urn:schemas-microsoft-com:vml", VML_OFFICE = "urn:schemas-microsoft-com:office:office", VML_EXCEL = "urn:schemas-microsoft-com:office:excel";
export interface NativeVmlNote {
  readonly cell: WorkbookCoordinate;
  readonly anchor: WorkbookDrawingAnchor;
  readonly visible: boolean;
  readonly locked: boolean;
  readonly printable: boolean;
  readonly lockText: boolean;
  readonly defaultSize: boolean;
  readonly disabled: boolean;
  readonly autoFill: boolean;
  readonly autoLine: boolean;
  readonly autoScale: boolean;
  readonly justifyLastLine: boolean;
  readonly textHorizontal?: string;
  readonly textVertical?: string;
  /** The appearance adapter must consume these remaining sources. */
  readonly shape: XmlNode;
  readonly shapeType?: XmlNode;
  readonly shapeStyle: ReadonlyMap<string, string>;
}
export function vmlBoolean(source: string | undefined, node: XmlNode, fallback: boolean): boolean {
  if (source === undefined) return fallback;
  if (["t", "true", "True", ""].includes(source)) return true;
  if (["f", "false", "False"].includes(source)) return false;
  return fail(node, "A VML boolean is invalid.");
}
export function vmlAttributes(node: XmlNode, names: readonly string[], extras: readonly [string, string][] = []): void {
  const allowed = new Set([...names.map(name => " " + name), ...extras.map(([uri, name]) => uri + " " + name)]);
  for (const [key] of node.attrs) if (!key.startsWith("http://www.w3.org/2000/xmlns/ ") && !allowed.has(key)) unsupported(node, "This VML note attribute is not supported.");
}
export function vmlStyle(node: XmlNode): ReadonlyMap<string, string> {
  const result = new Map<string, string>();
  for (const declaration of (value(node, "style") ?? "").split(";")) {
    if (!declaration.trim()) continue;
    const split = declaration.indexOf(":"), name = declaration.slice(0, split).trim().toLowerCase(), content = declaration.slice(split + 1).trim();
    if (split < 1 || !/^[a-z][a-z-]*$/.test(name) || !content || result.has(name)) fail(node, "VML note style declarations must be unambiguous.");
    result.set(name, content);
  }
  return result;
}
function leaf(node: XmlNode): string {
  attributes(node, []);
  if (node.content.some(item => typeof item !== "string")) fail(node, "A VML note property must contain only text.");
  return text(node).trim();
}
function clientData(node: XmlNode, shape: XmlNode, shapeType: XmlNode | undefined): NativeVmlNote {
  attributes(node, ["ObjectType"]);
  if (value(node, "ObjectType") !== "Note") unsupported(node, "Worksheet controls and other VML objects are outside the note boundary.");
  const flags = ["MoveWithCells", "SizeWithCells", "Visible", "Locked", "PrintObject", "LockText", "DefaultSize", "Disabled", "AutoFill", "AutoLine", "AutoScale", "JustLastX", "RowHidden", "ColHidden"];
  check(node, [...flags, "Anchor", "Row", "Column", "TextHAlign", "TextVAlign"], VML_EXCEL);
  const flag = (name: string, fallback: boolean): boolean => { const source = one(node, name, VML_EXCEL); return source ? vmlBoolean(leaf(source), source, fallback) : fallback; };
  const index = (name: string, maximum: number): number => {
    const source = one(node, name, VML_EXCEL); if (!source) fail(node, "A VML note is missing its cell address.");
    const raw = leaf(source), number = Number(raw);
    if (!/^[+]?\d+$/.test(raw) || !Number.isSafeInteger(number) || number > maximum) fail(source, "A VML note cell address is outside worksheet bounds.");
    return number;
  };
  const cell = { row: index("Row", 1048575), column: index("Column", 16383) }, anchor = one(node, "Anchor", VML_EXCEL);
  if (!anchor) unsupported(node, "A VML note without a cell anchor is not supported by this adapter.");
  const style = vmlStyle(shape), visibility = style.get("visibility"), explicitVisible = one(node, "Visible", VML_EXCEL);
  if (visibility !== undefined && !["visible", "hidden"].includes(visibility)) unsupported(shape, "This note visibility style is not supported.");
  if (explicitVisible && visibility !== undefined && flag("Visible", false) !== (visibility === "visible")) fail(shape, "VML note visibility declarations disagree.");
  const alignment = (name: string, allowed: readonly string[]): string | undefined => {
    const source = one(node, name, VML_EXCEL); if (!source) return undefined;
    const result = leaf(source).toLowerCase(); if (!allowed.includes(result)) unsupported(source, "This note text alignment is not supported."); return result;
  };
  // These two hints describe current row/column layout, which the worksheet
  // dimensions already own. Validate their lexical form without a second
  // authoritative hidden-row model.
  flag("RowHidden", false); flag("ColHidden", false);
  return { cell, anchor: readVmlAnchor(anchor, flag("MoveWithCells", false), flag("SizeWithCells", false)),
    visible: explicitVisible ? flag("Visible", false) : visibility === "visible", locked: flag("Locked", true), printable: flag("PrintObject", true),
    lockText: flag("LockText", true), defaultSize: flag("DefaultSize", true), disabled: flag("Disabled", false), autoFill: flag("AutoFill", true),
    autoLine: flag("AutoLine", true), autoScale: flag("AutoScale", false), justifyLastLine: flag("JustLastX", false),
    ...(one(node, "TextHAlign", VML_EXCEL) ? { textHorizontal: alignment("TextHAlign", ["left", "center", "right", "justify", "distributed"])! } : {}),
    ...(one(node, "TextVAlign", VML_EXCEL) ? { textVertical: alignment("TextVAlign", ["top", "center", "bottom", "justify", "distributed"])! } : {}),
    shape, ...(shapeType ? { shapeType } : {}), shapeStyle: style };
}
/** Read the explicit legacyDrawing relationship and note cells, retaining the
 * shape/type sources until the dedicated appearance mapping has consumed them. */
export function readVmlNoteRecords(sheet: XmlNode, opc: OpcPackage, maxItems: number): NativeVmlNote[] {
  const reference = one(sheet, "legacyDrawing"), related = opc.related(sheet.part, "vmlDrawing");
  if (!reference) {
    if (related) fail(sheet, "A VML drawing relationship has no worksheet reference.");
    return [];
  }
  attributes(reference, ["id"], R); check(reference, []);
  const id = attr(reference, "id", R), relation = id === undefined ? undefined : opc.relationships(sheet.part).get(id);
  if (!relation || relation.external || relation.type !== R + "/vmlDrawing" || relation.target !== related) fail(reference, "A legacy drawing reference is invalid or ambiguous.");
  const root = opc.read(relation.target);
  if (root.name !== "xml" || root.uri !== "") fail(root, "A VML note part has an invalid root.");
  attributes(root, []);
  const types = new Map<string, XmlNode>(), ids = new Set<string>(), notes: NativeVmlNote[] = [], cells = new Set<number>();
  let items = 0;
  const count = (): void => { if (++items > maxItems) throw new DocumentError("ResourceLimitExceeded", "VML notes exceed their record budget."); };
  for (const node of root.content) {
    if (typeof node === "string") { if (node.trim()) fail(root, "Unexpected text in a VML note part."); continue; }
    count();
    if (node.uri === VML && node.name === "shapetype") {
      const id = value(node, "id"); if (!id || ids.has(id)) fail(node, "A VML type identity is missing or duplicated."); ids.add(id); types.set(id, node);
    } else if (node.uri === VML_OFFICE && node.name === "shapelayout") {
      vmlAttributes(node, [], [[VML, "ext"]]);
      if (attr(node, "ext", VML) !== undefined && attr(node, "ext", VML) !== "edit") unsupported(node);
      check(node, ["idmap"], VML_OFFICE);
      for (const map of elements(node, "idmap", VML_OFFICE)) {
        vmlAttributes(map, ["data"], [[VML, "ext"]]); check(map, [], VML_OFFICE);
        if (attr(map, "ext", VML) !== undefined && attr(map, "ext", VML) !== "edit" || !/^\d+(?:,\d+)*$/.test(value(map, "data") ?? "")) fail(map, "A VML identity map is invalid.");
      }
    } else if (node.uri !== VML || node.name !== "shape") unsupported(node, "Only classic note shapes are supported in a worksheet VML part.");
  }
  for (const shape of elements(root, "shape", VML)) {
    const id = value(shape, "id"), reference = value(shape, "type"), type = reference?.startsWith("#") ? types.get(reference.slice(1)) : undefined;
    if (!id || ids.has(id) || reference !== undefined && !type) fail(shape, "A VML note shape identity or type reference is invalid."); ids.add(id);
    const client = one(shape, "ClientData", VML_EXCEL); if (!client) unsupported(shape, "A VML shape has no note client data.");
    const note = clientData(client, shape, type), key = note.cell.row * 16384 + note.cell.column;
    if (cells.has(key)) fail(shape, "A cell has multiple VML note shapes."); cells.add(key); notes.push(note);
  }
  return notes;
}
