import { dailyCi } from '../../../tests/ci-mode.js';
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS, createWorkbook, openDocument } from "@doffi/core";
import type { WorkbookSheet } from "@doffi/core";
import { openOpc } from "../src/ooxml/package.js";
import { readWorkbook } from "../src/xlsx/workbook.js";
import { readSheet, rowDimension } from "../src/xlsx/sheet.js";
import { scanWorksheet } from "../src/xlsx/scan.js";
import { ImportControl } from "../src/import-control.js";
import { S, xml } from "../src/ooxml/xml.js";

const limits = { ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 32 * 1024 * 1024, maxPackageBytes: 128 * 1024 * 1024, maxNodes: 2_000_000 };
const styles = { defaultCell: () => ({ cssClass: "cell-default", numberFormatId: "general" }), cell: () => ({ cssClass: "cell-default", numberFormatId: "general" }), tabStyle: () => "cell-default" };
const encoder = new TextEncoder(), decoder = new TextDecoder();
async function read(filename: string) {
  const opc = openOpc(readFileSync(filename), { limits }), source = readWorkbook(opc), sheets: WorkbookSheet[] = []; let cells = 0;
  for (const sheet of source.sheets) {
    const rows: WorkbookSheet["rowDimensions"] = [];
    const root = await scanWorksheet(opc.files.get(sheet.part)!, sheet.part, { limits, budget: opc.xmlBudget, control: new ImportControl(), maxRetainedNodes: 5000, sheetId: sheet.sheetId, pass: 0,
      cell() { cells++; }, row(node, row) { const dimension = rowDimension(node, row, styles); if (dimension) rows.push(dimension); },
    });
    sheets.push(readSheet(root, sheet, rows, opc, styles));
  }
  const files = new Map(createWorkbook({ documentId: "native-sheet" })), metadata = JSON.parse(decoder.decode(files.get("metadata.json")!));
  Object.assign(metadata.workbook, { sheets, activeSheetId: source.activeSheetId }); files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  return { model: openDocument(files, { kind: "workbook" }), cells };
}
describe("streaming native sheet geometry", () => {
  it.each(["examples/xlsx/03-structure.xlsx", "examples/xlsx/producer/03-structure.xlsx"])("preserves frozen/hidden dimensions and cross-chunk merges in %s", async filename => {
    const { model } = await read(filename), sheet = model.snapshot().metadata.workbook.sheets[0]!;
    expect(sheet.frozen).toEqual({ rows: 2, columns: 1 }); expect(sheet.columnDimensions[1]).toMatchObject({ from: 3, to: 3, width: 10, hidden: true });
    expect(sheet.rowDimensions.some(row => row.heightPt === 30)).toBe(true);
    expect(sheet.merges).toContainEqual({ from: { row: 255, column: 0 }, to: { row: 257, column: 0 } });
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
  });
  it("keeps hyperlink targets as inert metadata", async () => {
    const small = await read("examples/xlsx/01-values-text.xlsx"), sheet = small.model.snapshot().metadata.workbook.sheets[0]!;
    expect(sheet.hyperlinks?.[0]).toEqual({ range: { from: { row: 1, column: 0 }, to: { row: 1, column: 0 } }, url: "https://example.com/authorized-link", display: "文档链接", tooltip: "外部链接，不自动访问" });
  });
  it.skipIf(dailyCi)("streams the full real-data worksheet @scale", async () => {
    const large = await read("examples/xlsx/15-ten-thousand-rows.xlsx"); expect(large.cells).toBe(200000);
    expect(large.model.snapshot().metadata.workbook.sheets[0]!.rowDimensions).toEqual([]);
  });
  it("distinguishes a fixed sheet default, an explicit automatic row and a native automatic-height cache", () => {
    const parse = (source: string) => xml(encoder.encode(source), "sheet.xml", limits, { nodes: 0 });
    expect(rowDimension(parse(`<row xmlns="${S}" r="2" ht="80" customHeight="0"/>`), 1, styles)).toEqual({ from: 1, to: 1, autoHeight: true });
    expect(rowDimension(parse(`<row xmlns="${S}" r="2" ht="12" customHeight="1"/>`), 1, styles)).toEqual({ from: 1, to: 1, heightPt: 12 });
    expect(rowDimension(parse(`<row xmlns="${S}" r="2" ht="80"/>`), 1, styles)).toBeUndefined();
    const opc = openOpc(readFileSync("examples/xlsx/01-values-text.xlsx"), { limits }), source = readWorkbook(opc).sheets[0]!;
    const sheet = readSheet(parse(`<worksheet xmlns="${S}"><sheetFormatPr defaultRowHeight="18" customHeight="1"/><sheetData/></worksheet>`), source, [{ from: 1, to: 1, autoHeight: true }], opc, styles);
    expect(sheet.defaultRowHeightCustom).toBe(true); expect(sheet.defaultRowHeightPt).toBe(18); expect(sheet.rowDimensions[0]?.autoHeight).toBe(true);
  });
  it("infers omitted addresses and rejects duplicate, misplaced, unordered and out-of-bounds cells", async () => {
    const scan = async (content: string) => {
      const coordinates: unknown[] = [];
      await scanWorksheet(encoder.encode(`<worksheet xmlns="${S}">${content}</worksheet>`), "sheet.xml", { limits, budget: { nodes: 0 }, control: new ImportControl(), maxRetainedNodes: 100, sheetId: "s", pass: 0, cell(_, coordinate) { coordinates.push(coordinate); } });
      return coordinates;
    };
    expect(await scan('<sheetData><row><c/><c r="D1"/><c/></row><row r="5"><c/></row></sheetData>')).toEqual([{ row: 0, column: 0 }, { row: 0, column: 3 }, { row: 0, column: 4 }, { row: 4, column: 0 }]);
    for (const content of ['<sheetData><row><c r="A1"/><c r="A1"/></row></sheetData>', '<sheetData><row><c r="A2"/></row></sheetData>', '<sheetData><row><c r="XFE1"/></row></sheetData>', '<sheetData><row r="2"/><row r="1"/></sheetData>', '<sheetData><c r="A1"/></sheetData>', '<sheetData/><sheetData/>', '']) await expect(scan(content)).rejects.toMatchObject({ code: "InvalidDocument" });
  });
});
