import { attr, check, child, elements, fail, R, S, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { OpcPackage } from "../ooxml/package.js";
import { Context, integer, rootName, unsigned, value } from "./context.js";

export type Cell = string | number | null;
export interface ConstantWorkbook { read(formula: string, source: XmlNode): Cell[] }
function address(value: string, source: XmlNode): { column: number; row: number } {
  const match = /^\$?([A-Z]{1,3})\$?([1-9]\d*)$/i.exec(value);
  if (!match) return unsupported(source, "Only ordinary bounded worksheet range references are supported.");
  let column = 0;
  for (const character of match[1]!.toUpperCase()) column = column * 26 + character.charCodeAt(0) - 64;
  const row = Number(match[2]);
  if (column > 16384 || row > 1048576) fail(source, "A chart range exceeds worksheet bounds.");
  return { column, row };
}
function related(opc: OpcPackage, source: string, id: string | undefined, kind: string): string {
  const relation = id && opc.relationships(source).get(id);
  if (!relation || relation.type !== R + "/" + kind) fail({ part: source, name: "Relationship" }, "A chart workbook relationship is missing or invalid.");
  if (relation.external) unsupported({ part: source, name: "Relationship" } as XmlNode, "External workbook dependencies are not supported.");
  return relation.target;
}
export function constants(context: Context, path: string): ConstantWorkbook {
  const opc = context.workbook(path), rootRelation = [...opc.relationships("").values()].filter(r => r.type === R + "/officeDocument");
  if (rootRelation.length !== 1) fail({ part: path, name: "workbook" }, "An embedded workbook requires one main part.");
  const main = related(opc, "", rootRelation[0]!.id, "officeDocument");
  if (opc.contentType(main) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml") unsupported({ part: path, name: "workbook" } as XmlNode, "Only a plain XLSX constant workbook is supported.");
  const source = opc.read(main); rootName(source, "workbook", S);
  check(source, ["fileVersion", "workbookPr", "bookViews", "sheets", "calcPr", "extLst"], S);
  const strings: string[] = [];
  const stringsParts = [...opc.relationships(main).values()].filter(r => r.type === R + "/sharedStrings");
  if (stringsParts.length > 1) fail(source, "Shared strings are ambiguous.");
  if (stringsParts[0]) {
    const part = related(opc, main, stringsParts[0].id, "sharedStrings"), sst = opc.read(part); rootName(sst, "sst", S);
    for (const item of elements(sst, "si", S)) { context.count(); strings.push(elements(item, "t", S).map(text).join("") + elements(item, "r", S).map(r => text(child(r, "t", S))).join("")); }
  }
  // Every worksheet is checked for formulas/external dependencies, including
  // cells outside the selected ranges. The model retains only chart constants.
  const sheets = new Map<string, Map<string, Cell>>();
  for (const entry of elements(child(source, "sheets", S), "sheet", S)) {
    const name = value(entry, "name"), part = related(opc, main, attr(entry, "id", R), "worksheet");
    if (!name || sheets.has(name)) fail(entry, "Worksheet names are missing or duplicated.");
    const sheet = opc.read(part); rootName(sheet, "worksheet", S);
    if (child(sheet, "externalLinks", S) || child(sheet, "tableParts", S)) unsupported(sheet, "Linked workbook data is not supported.");
    const cells = new Map<string, Cell>();
    for (const row of elements(child(sheet, "sheetData", S), "row", S)) for (const cell of elements(row, "c", S)) {
      context.count();
      const coordinate = value(cell, "r");
      if (!coordinate) fail(cell, "An embedded cell address is missing.");
      const position = address(coordinate, cell), key = position.row + ":" + position.column;
      if (cells.has(key)) fail(cell, "An embedded cell address is duplicated.");
      if (child(cell, "f", S)) unsupported(cell, "Cell formulas are outside the constant chart data scope.");
      const type = value(cell, "t"), v = child(cell, "v", S);
      let result: Cell;
      if (type === "inlineStr") { const inline = child(cell, "is", S); result = elements(inline, "t", S).map(text).join("") + elements(inline, "r", S).map(r => text(child(r, "t", S))).join(""); }
      else if (type === "s") { const index = unsigned({ ...cell, attrs: new Map([[" val", text(v)]]) }, "val"); result = strings[index] ?? fail(cell, "A shared string reference is invalid."); }
      else if (type === "str" || type === "d") result = text(v);
      else if (type === undefined || type === "n") {
        const raw = text(v); result = v ? Number(raw) : null;
        if (v && (!raw.trim() || !Number.isFinite(result))) fail(cell, "An embedded numeric constant is invalid.");
      } else unsupported(cell, "Only numeric, string and blank chart constants are supported.");
      cells.set(key, result);
    }
    for (const relation of opc.relationships(part).values()) if (relation.external) unsupported(sheet, "External worksheet dependencies are not supported.");
    sheets.set(name, cells);
  }
  for (const relation of opc.relationships(main).values()) if (relation.external || /\/(externalLink|vbaProject|connections)$/.test(relation.type)) unsupported(source, "External or executable chart workbook dependencies are not supported.");
  return { read(formula, node) {
    const match = /^(?:'((?:[^']|'')+)'|([^'!\[\]():]+))!([^!]+)$/.exec(formula.trim());
    if (!match) unsupported(node, "Chart formulas must be local ordinary ranges.");
    const sheet = sheets.get((match[1]?.replaceAll("''", "'") ?? match[2])!);
    if (!sheet) fail(node, "A chart range refers to an undefined worksheet.");
    const parts = match[3]!.split(":"), start = address(parts[0]!, node), end = address(parts[1] ?? parts[0]!, node);
    if (parts.length > 2 || start.column > end.column || start.row > end.row || start.column !== end.column && start.row !== end.row) unsupported(node, "A chart vector must be one contiguous row or column.");
    const length = end.row - start.row + end.column - start.column + 1;
    context.count(length);
    const values: Cell[] = [];
    for (let index = 0; index < length; index++) values.push(sheet.get((start.row + (start.row === end.row ? 0 : index)) + ":" + (start.column + (start.column === end.column ? 0 : index))) ?? null);
    return values;
  } };
}

export function cachedVector(context: Context, parent: XmlNode, numeric: boolean, workbook?: ConstantWorkbook): Cell[] {
  const ref = child(parent, numeric ? "numRef" : "strRef", "http://schemas.openxmlformats.org/drawingml/2006/chart");
  const literal = child(parent, numeric ? "numLit" : "strLit", "http://schemas.openxmlformats.org/drawingml/2006/chart");
  const c = "http://schemas.openxmlformats.org/drawingml/2006/chart";
  if (!!ref === !!literal) unsupported(parent, "A chart vector requires one supported constant or cached reference.");
  const cache = literal ?? child(ref, numeric ? "numCache" : "strCache", c);
  if (!cache) fail(parent, "A chart reference cache is missing.");
  check(cache, numeric ? ["formatCode", "ptCount", "pt", "extLst"] : ["ptCount", "pt", "extLst"], c);
  const count = child(cache, "ptCount", c);
  if (!count) fail(cache, "A chart cache length is missing.");
  const length = integer(count, "val");
  if (length < 0) fail(count, "A chart cache length cannot be negative.");
  context.count(length);
  const values = Array<Cell>(length).fill(numeric ? null : ""), indices = new Set<number>();
  for (const point of elements(cache, "pt", c)) {
    const index = integer(point, "idx"), raw = child(point, "v", c);
    if (index < 0 || index >= length || indices.has(index) || !raw) fail(point, "Chart cache indices are invalid or duplicated.");
    indices.add(index);
    if (numeric) { const source = text(raw); if (!source.trim() || !Number.isFinite(Number(source))) fail(raw, "A chart cache number is invalid."); values[index] = Number(source); }
    else values[index] = text(raw);
  }
  if (ref) {
    const formula = text(child(ref, "f", c));
    if (!formula) fail(ref, "A chart range reference is missing.");
    if (workbook) {
      const cells = workbook.read(formula, ref);
      if (cells.length !== values.length || cells.some((cell, index) => numeric ? cell !== values[index] : String(cell ?? "") !== values[index])) fail(ref, "Chart cache and embedded constant values disagree.");
    } else if (!/^(?:label \d+|categories|\d+)$/.test(formula) && !/^(?:'[^\[\]']*(?:''[^\[\]']*)*'|[^'!\[\]():]+)!\$?[A-Z]{1,3}\$?[1-9]\d*(?::\$?[A-Z]{1,3}\$?[1-9]\d*)?$/i.test(formula)) unsupported(ref, "A cache-only chart cannot evaluate formulas or use external workbook references.");
  }
  return values;
}
