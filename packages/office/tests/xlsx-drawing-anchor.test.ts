import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS, WorkbookDrawingAnchorSchema } from "@doffi/core";
import { openOpc } from "../src/ooxml/package.js";
import { xml } from "../src/ooxml/xml.js";
import { readDrawingAnchor, readNoteObjectAnchor, readVmlAnchor, XDR } from "../src/xlsx/drawing-anchor.js";

const encoder = new TextEncoder();
function parse(source: string) { return xml(encoder.encode(source), "drawing.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 }); }
const marker = (kind: string, col = 1, row = 2, x = "12700", y = "25400") => `<${kind}><col>${col}</col><colOff>${x}</colOff><row>${row}</row><rowOff>${y}</rowOff></${kind}>`;
const anchor = (source: string) => readDrawingAnchor(parse(source));
describe("workbook drawing and note anchor primitives", () => {
  it("reads all independent picture/shape/textbox and three-chart anchors and native producer variants", () => {
    let count = 0;
    for (const directory of ["", "producer/"]) for (const name of ["07-three-charts", "08-images-shapes"]) {
      const opc = openOpc(readFileSync(`examples/xlsx/${directory}${name}.xlsx`), {});
      for (const path of opc.files.keys()) if (/^xl\/drawings\/[^/]+\.xml$/.test(path)) {
        const root = opc.read(path);
        for (const node of root.content) if (typeof node !== "string") {
          const record = readDrawingAnchor(node); expect(WorkbookDrawingAnchorSchema.safeParse(record.anchor).success).toBe(true);
          expect(["pic", "sp", "graphicFrame"]).toContain(record.object.name); expect(record.locked).toBe(true); expect(record.printable).toBe(true); count++;
        }
      }
    }
    expect(count).toBe(12);
  });
  it("preserves three native response modes and offsets exceeding a cell instead of normalizing against a viewport", () => {
    for (const [editAs, moveWithCells, sizeWithCells] of [["twoCell", true, true], ["oneCell", true, false], ["absolute", false, false]] as const) {
      const result = anchor(`<twoCellAnchor xmlns="${XDR}" editAs="${editAs}">${marker("from", 1, 0, "-12700", "380880")}${marker("to", 3, 8)}<sp/><clientData fLocksWithSheet="0" fPrintsWithSheet="0"/></twoCellAnchor>`);
      expect(result).toMatchObject({ anchor: { kind: "two-cell", from: { row: 0, column: 1, offsetXPt: -1, offsetYPt: 380880 / 12700 }, moveWithCells, sizeWithCells }, locked: false, printable: false });
    }
  });
  it("converts EMU and universal measures, retaining valid zero extents and fixed document geometry", () => {
    const result = anchor(`<absoluteAnchor xmlns="${XDR}"><pos x="1in" y="-2.54cm"/><ext cx="0" cy="12700"/><sp/><clientData/></absoluteAnchor>`);
    expect(result.anchor).toEqual({ kind: "absolute", xPt: 72, yPt: -72, widthPt: 0, heightPt: 1 });
    expect(WorkbookDrawingAnchorSchema.safeParse(result.anchor).success).toBe(true);
    const one = anchor(`<oneCellAnchor xmlns="${XDR}">${marker("from")}<ext cx="2032000" cy="1524000"/><pic/><clientData/></oneCellAnchor>`);
    expect(one.anchor).toMatchObject({ kind: "one-cell", widthPt: 160, heightPt: 120 });
  });
  it("maps VML pixels and standard commentPr markers to the same point-unit source", () => {
    const vml = readVmlAnchor(parse('<Anchor xmlns="urn:schemas-microsoft-com:office:excel">2,10,1,2,5,20,5,4</Anchor>'), true, true);
    expect(vml).toEqual({ kind: "two-cell", from: { row: 1, column: 2, offsetXPt: 7.5, offsetYPt: 1.5 }, to: { row: 5, column: 5, offsetXPt: 15, offsetYPt: 3 }, moveWithCells: true, sizeWithCells: true });
    const native = readNoteObjectAnchor(parse(`<anchor xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" moveWithCells="1"><from xmlns="${XDR}"><col>2</col><colOff>95250</colOff><row>1</row><rowOff>19050</rowOff></from><to xmlns="${XDR}"><col>5</col><colOff>190500</colOff><row>5</row><rowOff>38100</rowOff></to></anchor>`));
    expect(native).toEqual({ ...vml, sizeWithCells: false });
  });
  it.each([
    '<absoluteAnchor><pos x="0" y="0"/><ext cx="-1" cy="10"/><sp/><clientData/></absoluteAnchor>',
    '<absoluteAnchor><pos x="0" y="0"/><ext cx="1pt" cy="10"/><sp/><clientData/></absoluteAnchor>',
    '<absoluteAnchor><pos x="NaN" y="0"/><ext cx="1" cy="10"/><sp/><clientData/></absoluteAnchor>',
    '<absoluteAnchor><pos x="0" y="0"/><ext cx="1" cy="10"/><sp/><pic/><clientData/></absoluteAnchor>',
    '<absoluteAnchor><pos x="0" y="0"/><ext cx="1" cy="10"/><sp/></absoluteAnchor>',
    `<oneCellAnchor>${marker("from", 16384)}<ext cx="1" cy="1"/><pic/><clientData/></oneCellAnchor>`,
    `<twoCellAnchor editAs="unknown">${marker("from")}${marker("to")}<sp/><clientData/></twoCellAnchor>`,
  ])("rejects malformed or ambiguous geometry %#", source => {
    expect(() => anchor(source.replace(/^<(\w+)/, `<$1 xmlns="${XDR}"`))).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
  it("rejects malformed VML lists, out-of-sheet anchors and undeclared properties", () => {
    for (const source of ["1,2,3", "1,0,2,0,16384,0,4,0", "1,0,2,0,3,0,1048576,0", "1,NaN,2,0,3,0,4,0"]) expect(() => readVmlAnchor(parse(`<Anchor>${source}</Anchor>`), true, true)).toThrow();
    expect(() => anchor(`<absoluteAnchor xmlns="${XDR}" surprise="1"><pos x="0" y="0"/><ext cx="1" cy="1"/><sp/><clientData/></absoluteAnchor>`)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
});
