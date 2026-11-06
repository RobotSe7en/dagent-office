import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument, DEFAULT_DOCUMENT_READ_LIMITS, WorkbookDrawingAnchorSchema } from "@dagent-office/core";
import { openOpc, openOpcFiles } from "../src/ooxml/package.js";
import type { OpcPackage } from "../src/ooxml/package.js";
import { readWorkbook } from "../src/xlsx/workbook.js";
import { XlsxStyles } from "../src/xlsx/styles.js";
import { readClassicComments } from "../src/xlsx/classic-comments.js";
import { readNoteBoxes } from "../src/xlsx/note-boxes.js";
import { WorkbookContent } from "../src/xlsx/content.js";
import { mapWorkbookComments } from "../src/xlsx/comment-mapping.js";
import { readVmlNoteRecords } from "../src/xlsx/vml-notes.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const source = (producer = false) => openOpc(readFileSync(`examples/xlsx/${producer ? "producer/" : ""}09-classic-notes.xlsx`), {});
function read(opc: OpcPackage) { return readVmlNoteRecords(opc.read("xl/worksheets/sheet1.xml"), opc, 1000); }
function variant(edit: (source: string) => string) {
  const files = new Map(source().files), part = "xl/drawings/comments.vml";
  files.set(part, encoder.encode(edit(decoder.decode(files.get(part)!)))); return openOpcFiles(files, {});
}
function mapBoxes(opc: OpcPackage) {
  const workbook = readWorkbook(opc), styles = new XlsxStyles(opc, workbook.part, 1_000_000);
  const notes = readClassicComments(opc, workbook.sheets, styles, 1000, 1_000_000);
  const boxes = readNoteBoxes(opc, workbook.sheets.map(sheet => ({ sheetId: sheet.sheetId, root: opc.read(sheet.part) })), notes, styles, 1000);
  return { notes, boxes, styles };
}
describe("classic note VML geometry records", () => {
  it("maps solid note appearance and flags into one persisted source with exact rich-text bodies", () => {
    for (const producer of [false, true]) {
      const input = mapBoxes(source(producer)), content = new WorkbookContent(DEFAULT_DOCUMENT_READ_LIMITS);
      const mapped = mapWorkbookComments({ people: [], threads: [] }, [], input.notes, input.boxes, content, "notes");
      expect(mapped.comments[0]).toMatchObject({ format: "classic", box: { visible: false, autoFill: false, printable: true, anchor: { kind: "two-cell" } } });
      expect(input.styles.stylesheet()).toContain("--office-note-fill:#ffffe1"); expect(input.styles.stylesheet()).toContain("--office-note-stroke-width:0.75pt");
      expect(input.styles.stylesheet()).toContain("--office-note-shadow:black");
      const files = new Map(createWorkbook({ documentId: "notes" })), manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
      manifest.content = "content.html"; Object.assign(metadata.workbook, { stories: content.stories, commentAuthors: mapped.authors, comments: mapped.comments });
      for (const [path, value] of [["manifest.json", manifest], ["metadata.json", metadata], ["annotations.json", mapped.annotations]] as const) files.set(path, encoder.encode(JSON.stringify(value)));
      files.set("content.html", encoder.encode(content.html()!)); files.set("styles.css", encoder.encode(decoder.decode(files.get("styles.css")!) + input.styles.stylesheet()));
      const model = openDocument(files, { kind: "workbook" }); expect(model.getNode(mapped.annotations[0]!.bodyNodeId!)?.text).toBe("请确认：预算😀禰󠄀");
      expect(openDocument(model.serialize()).snapshot()).toEqual(model.snapshot());
      const comment = metadata.workbook.comments[0]; comment.box.cssClass = "missing"; files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
      expect(() => openDocument(files)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    }
  });
  it("does not ignore unknown shape appearance, non-note types, duplicate children or disconnected note shapes", () => {
    for (const edit of [
      (s: string) => s.replace('fillcolor="#ffffe1"', 'fillcolor="url(https://example.invalid/image)"'),
      (s: string) => s.replace('<v:fill color2=', '<v:fill type="gradient" color2='),
      (s: string) => s.replace('o:spt="202"', 'o:spt="1"'),
      (s: string) => s.replace('visibility:hidden', 'visibility:hidden;rotation:45'),
      (s: string) => s.replace('<v:shadow on=', '<v:shadow/><v:shadow on='),
      (s: string) => s.replace('<x:Column>1</x:Column>', '<x:Column>2</x:Column>'),
    ]) expect(() => mapBoxes(variant(edit))).toThrow();
  });
  it("reads standard commentPr anchors without a competing VML representation", () => {
    const files = new Map(source().files), sheetPart = "xl/worksheets/sheet1.xml", commentsPart = "xl/comments/comment1.xml";
    files.set(sheetPart, encoder.encode(decoder.decode(files.get(sheetPart)!).replace(/<legacyDrawing [^>]*\/>/, '')));
    const rels = "xl/worksheets/_rels/sheet1.xml.rels";
    files.set(rels, encoder.encode(decoder.decode(files.get(rels)!).replace(/<Relationship [^>]*Type="[^"]*\/vmlDrawing"[^>]*\/>/, '')));
    const anchor = '<anchor moveWithCells="1" sizeWithCells="0" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"><xdr:from><xdr:col>2</xdr:col><xdr:colOff>95250</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>19050</xdr:rowOff></xdr:from><xdr:to><xdr:col>5</xdr:col><xdr:colOff>190500</xdr:colOff><xdr:row>5</xdr:row><xdr:rowOff>38100</xdr:rowOff></xdr:to></anchor>';
    files.set(commentsPart, encoder.encode(decoder.decode(files.get(commentsPart)!).replace('</comment>', `<commentPr altText="预算说明" print="0" textHAlign="center">${anchor}</commentPr></comment>`)));
    const mapped = mapBoxes(openOpcFiles(files, {}));
    expect([...mapped.boxes.values()][0]).toMatchObject({ source: "commentPr", box: { altText: "预算说明", printable: false, visible: false, anchor: { kind: "two-cell", moveWithCells: true, sizeWithCells: false } } });
  });
  it("preserves source and Calc cell anchors, visibility and independent row/column response flags", () => {
    for (const producer of [false, true]) {
      const records = read(source(producer)); expect(records).toHaveLength(1);
      expect(records[0]).toMatchObject({ cell: { row: 1, column: 1 }, visible: false, locked: true, lockText: true, printable: true, autoFill: false,
        anchor: { kind: "two-cell", from: { row: 1, column: 2, offsetXPt: 7.5, offsetYPt: 1.5 }, to: { row: 5, column: 5, offsetXPt: 15, offsetYPt: 3 }, moveWithCells: true, sizeWithCells: true } });
      expect(WorkbookDrawingAnchorSchema.safeParse(records[0]!.anchor).success).toBe(true);
      expect(records[0]!.shapeType?.name).toBe("shapetype");
      if (producer) expect(records[0]).toMatchObject({ textHorizontal: "left", textVertical: "top" });
    }
  });
  it("takes geometry from cell markers without substituting native CSS position caches", () => {
    const base = read(source())[0]!, changed = read(variant(s => s.replace('margin-left:80pt;margin-top:30pt;width:150pt;height:75pt', 'margin-left:600pt;margin-top:700pt;width:20pt;height:10pt')))[0]!;
    expect(changed.anchor).toEqual(base.anchor); expect(changed.shapeStyle.get("margin-left")).toBe("600pt");
    const flags = read(variant(s => s.replace('<x:MoveWithCells/>', '<x:MoveWithCells>False</x:MoveWithCells>').replace('<x:SizeWithCells/>', '<x:SizeWithCells>False</x:SizeWithCells>').replace('visibility:hidden', 'visibility:visible')))[0]!;
    expect(flags).toMatchObject({ visible: true, anchor: { moveWithCells: false, sizeWithCells: false } });
  });
  it.each([
    ["contradictory visibility", (s: string) => s.replace('<x:Row>', '<x:Visible/><x:Row>')],
    ["unknown parent type", (s: string) => s.replace('type="#_x0000_t202"', 'type="#missing"')],
    ["invalid cell row", (s: string) => s.replace('<x:Row>1</x:Row>', '<x:Row>1048576</x:Row>')],
    ["missing cell column", (s: string) => s.replace('<x:Column>1</x:Column>', '')],
    ["repeated style", (s: string) => s.replace('visibility:hidden', 'visibility:hidden;visibility:visible')],
    ["invalid flag", (s: string) => s.replace('<x:AutoFill>False</x:AutoFill>', '<x:AutoFill>maybe</x:AutoFill>')],
  ])("rejects %s", (_name, edit) => {
    expect(() => read(variant(edit))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
  it("rejects controls, unknown client data and a dangling legacy drawing relationship", () => {
    expect(() => read(variant(s => s.replace('ObjectType="Note"', 'ObjectType="Button"')))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(() => read(variant(s => s.replace('</x:ClientData>', '<x:FmlaMacro>run()</x:FmlaMacro></x:ClientData>')))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    const files = new Map(source().files), part = "xl/worksheets/sheet1.xml";
    files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace(/<legacyDrawing [^>]*\/>/, '')));
    expect(() => read(openOpcFiles(files, {}))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
});
