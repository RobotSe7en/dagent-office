import type { WorkbookNoteBox } from "@dagent-office/core";
import { check, elements, fail, S, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import type { OpcPackage } from "../ooxml/package.js";
import type { NativeClassicNotes } from "./classic-comments.js";
import { readNoteObjectAnchor } from "./drawing-anchor.js";
import { attributes, boolean, one, spreadsheetText, value } from "./values.js";
import { readVmlNoteRecords, VML } from "./vml-notes.js";
import { validateVmlNoteType, vmlNoteAppearance } from "./vml-note-appearance.js";
import type { XlsxStyles } from "./styles.js";

const key = (sheetId: string, row: number, column: number) => JSON.stringify([sheetId, row, column]);
function properties(node: XmlNode, styles: XlsxStyles): WorkbookNoteBox {
  attributes(node, ["locked", "defaultSize", "print", "disabled", "autoFill", "autoLine", "altText", "textHAlign", "textVAlign", "lockText", "justLastX", "autoScale"]); check(node, ["anchor"], S);
  const anchor = one(node, "anchor"); if (!anchor) fail(node, "A native note property record is missing its anchor.");
  const horizontal = value(node, "textHAlign") ?? "left", vertical = value(node, "textVAlign") ?? "top";
  if (!["left", "center", "right", "justify", "distributed"].includes(horizontal) || !["top", "center", "bottom", "justify", "distributed"].includes(vertical)) unsupported(node, "This note text alignment is not supported.");
  const cssClass = styles.objectStyle({ "text-align": horizontal === "distributed" ? "justify" : horizontal, "--office-note-horizontal": horizontal, "--office-note-vertical": vertical })!;
  const alt = value(node, "altText");
  return { anchor: readNoteObjectAnchor(anchor), cssClass, visible: false, printable: boolean(node, "print", true), locked: boolean(node, "locked", true),
    lockText: boolean(node, "lockText", true), defaultSize: boolean(node, "defaultSize", true), disabled: boolean(node, "disabled"), autoFill: boolean(node, "autoFill", true),
    autoLine: boolean(node, "autoLine", true), autoScale: boolean(node, "autoScale"), justifyLastLine: boolean(node, "justLastX"), ...(alt === undefined ? {} : { altText: spreadsheetText(alt, node) }) };
}
/** Bind each legacy shape to its actual note cell. A separate shape or an
 * unconsumed control must never disappear during a successful import. */
export function readNoteBoxes(opc: OpcPackage, sheets: readonly { sheetId: string; root: XmlNode }[], notes: NativeClassicNotes, styles: XlsxStyles, maxItems: number): ReadonlyMap<string, WorkbookNoteBox> {
  const noteCells = new Map(notes.notes.map(note => [key(note.sheetId, note.cell.row, note.cell.column), note])), result = new Map<string, WorkbookNoteBox>(), parts = new Set<string>();
  for (const sheet of sheets) {
    const part = opc.related(sheet.root.part, "vmlDrawing");
    if (part) {
      if (parts.has(part)) fail(sheet.root, "A VML note part is shared between worksheets."); parts.add(part);
      for (const type of elements(opc.read(part), "shapetype", VML)) validateVmlNoteType(type);
    }
    const records = readVmlNoteRecords(sheet.root, opc, maxItems - result.size);
    for (const record of records) {
      const position = key(sheet.sheetId, record.cell.row, record.cell.column);
      if (!noteCells.has(position)) fail(record.shape, "A VML note shape has no classic note or compatibility placeholder at its cell.");
      const cssClass = styles.objectStyle(vmlNoteAppearance(record))!, alt = value(record.shape, "alt");
      result.set(position, { anchor: record.anchor, cssClass, visible: record.visible, printable: record.printable, locked: record.locked, lockText: record.lockText,
        defaultSize: record.defaultSize, disabled: record.disabled, autoFill: record.autoFill, autoLine: record.autoLine, autoScale: record.autoScale, justifyLastLine: record.justifyLastLine,
        ...(alt === undefined ? {} : { altText: spreadsheetText(alt, record.shape) }) });
    }
  }
  for (const note of notes.notes) if (note.properties) {
    const position = key(note.sheetId, note.cell.row, note.cell.column), vml = result.get(position), native = properties(note.properties, styles);
    // Two independent native geometry sources cannot be silently reconciled.
    // A later format extension may define a precedence rule; until then make
    // the ambiguity explicit rather than dropping a supported property set.
    if (vml) unsupported(note.properties, "A note with both VML and commentPr geometry requires an explicit supported precedence mapping.");
    result.set(position, native);
  }
  return new Map(notes.notes.flatMap(note => { const box = result.get(key(note.sheetId, note.cell.row, note.cell.column)); return box ? [[`${note.sheetId}:${note.cell.row}:${note.cell.column}`, box] as const] : []; }));
}
