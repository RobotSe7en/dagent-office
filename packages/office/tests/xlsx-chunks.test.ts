import { dailyCi } from '../../../tests/ci-mode.js';
import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, createSharedWorkbookFormulas, DEFAULT_DOCUMENT_READ_LIMITS, openDocument } from "@doffi/core";
import { openOpc } from "../src/ooxml/package.js";
import { ImportControl } from "../src/import-control.js";
import { readCell, cellFormula } from "../src/xlsx/cell.js";
import { WorkbookChunks, utf8Bytes } from "../src/xlsx/chunks.js";
import { scanWorksheet } from "../src/xlsx/scan.js";
import { readSharedStrings } from "../src/xlsx/strings.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const cell = (column: number, value: string) => `<td data-column="${column}" data-value-type="text">${value}</td>`;
describe("bounded native cell-to-chunk assembly", () => {
  it("roundtrips independently generated one-millionth-row and cross-band cells with exact UTF-8 budgets", () => {
    const chunks = new WorkbookChunks(DEFAULT_DOCUMENT_READ_LIMITS); chunks.startSheet("sheet-1");
    for (const [row, column, value] of [[0, 0, "a😀禰󠄀"], [0, 32, "b"], [1, 0, "c"], [255, 0, "d"], [256, 0, "e"], [1_048_575, 16383, "far"]] as const) chunks.append({ row, column }, cell(column, value));
    chunks.finish();
    expect(chunks.descriptors.map(chunk => [chunk.row, chunk.column])).toEqual([[0, 0], [0, 32], [256, 0], [1_048_320, 16352]]);
    for (const [path, bytes] of createWorkbook({ documentId: "chunks" })) chunks.addFile(path, bytes);
    const manifest = JSON.parse(decoder.decode(chunks.files.get("manifest.json")!)); manifest.chunks = chunks.descriptors; chunks.files.set("manifest.json", encoder.encode(JSON.stringify(manifest)));
    const model = openDocument(chunks.files, { kind: "workbook" }); expect(model.getCell("sheet-1", 0, 0)?.value).toEqual({ type: "text", value: "a😀禰󠄀" });
    expect(model.getCell("sheet-1", 1_048_575, 16383)?.text).toBe("far");
    for (const descriptor of chunks.descriptors) expect(createHash("sha256").update(chunks.files.get(descriptor.path)!).digest("hex")).toBe(descriptor.sha256);
    expect(utf8Bytes("a😀禰󠄀é\r\n")).toBe(encoder.encode("a😀禰󠄀é\r\n").length);
  });
  it.skipIf(dailyCi)("maps all 200,000 actual native cell records to 40 chunks and reopens typed formula caches @scale", async () => {
    const nativeLimits = { ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 32 * 1024 * 1024, maxNodes: 2_000_000 }, control = new ImportControl();
    const opc = openOpc(readFileSync("examples/xlsx/15-ten-thousand-rows.xlsx"), { limits: nativeLimits }), part = "xl/worksheets/sheet1.xml", shared = createSharedWorkbookFormulas();
    const styles = { defaultCell: () => ({ cssClass: "cell-default", numberFormatId: "general" }), run: () => undefined, cell: () => ({ cssClass: "cell-default", numberFormatId: "general" }), tabStyle: () => "cell-default" };
    const strings = await readSharedStrings(undefined, "missing.xml", styles, { limits: nativeLimits, budget: opc.xmlBudget, maxRetainedNodes: 1000, maxStrings: 250000, maxCodeUnits: 16 * 1024 * 1024, control });
    const chunks = new WorkbookChunks(DEFAULT_DOCUMENT_READ_LIMITS); chunks.startSheet("sheet-1"); let cells = 0, formulas = 0;
    const scan = { limits: nativeLimits, budget: opc.xmlBudget, maxRetainedNodes: 1000, control, sheetId: "sheet-1" };
    await scanWorksheet(opc.files.get(part)!, part, { ...scan, pass: 0, cell(node, coordinate) { cellFormula(node, coordinate, shared, 0); } });
    await scanWorksheet(opc.files.get(part)!, part, { ...scan, pass: 1, cell(node, coordinate) { const source = readCell(node, coordinate, styles, strings, shared); chunks.append(coordinate, source.html); cells++; if (source.formula) formulas++; } });
    chunks.finish(); expect(cells).toBe(200000); expect(formulas).toBe(20000); expect(chunks.descriptors).toHaveLength(40);
    const files = new Map([...createWorkbook({ documentId: "native-real-data" }), ...chunks.files]), manifest = JSON.parse(decoder.decode(files.get("manifest.json")!));
    manifest.chunks = chunks.descriptors; files.set("manifest.json", encoder.encode(JSON.stringify(manifest)));
    const model = openDocument(files, { kind: "workbook" });
    expect(model.getCell("sheet-1", 9999, 18)?.value).toEqual({ type: "number", value: 3_200_152 });
    expect(model.getCell("sheet-1", 9999, 19)?.value).toEqual({ type: "number", value: 3_400_152 });
    expect(model.getCell("sheet-1", 9999, 18)?.formula).toEqual({ expression: "=SUM(C10000:R10000)", cacheState: "imported" });
  });
  it("enforces cumulative byte/file limits across worksheet boundaries and rejects invalid output positions", () => {
    const small = new WorkbookChunks({ ...DEFAULT_DOCUMENT_READ_LIMITS, maxFileBytes: 200 }); small.startSheet("s");
    expect(() => small.append({ row: 0, column: 0 }, cell(0, "x".repeat(100)))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const count = new WorkbookChunks({ ...DEFAULT_DOCUMENT_READ_LIMITS, maxFiles: 1 }); count.startSheet("first"); count.append({ row: 0, column: 0 }, cell(0, "x")); count.startSheet("second");
    expect(() => count.append({ row: 0, column: 0 }, cell(0, "y"))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const total = new WorkbookChunks({ ...DEFAULT_DOCUMENT_READ_LIMITS, maxPackageBytes: 400 }); total.startSheet("s1"); total.append({ row: 0, column: 0 }, cell(0, "x")); total.startSheet("s2");
    expect(() => total.append({ row: 0, column: 0 }, cell(0, "y"))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const positions = new WorkbookChunks(DEFAULT_DOCUMENT_READ_LIMITS); positions.startSheet("s"); positions.append({ row: 0, column: 0 }, cell(0, "x"));
    expect(() => positions.append({ row: 0, column: 0 }, cell(0, "x"))).toThrow(); expect(() => positions.append({ row: 0.5, column: 0 }, cell(0, "x"))).toThrow();
  });
});
