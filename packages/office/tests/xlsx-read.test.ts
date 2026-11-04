import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";
import { readFileSync, readdirSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import type { XlsxRecords } from "../src/xlsx/read.js";
import type { XlsxReadOptions } from "../src/xlsx/options.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const source = (name: string) => readFileSync(`examples/xlsx/${name}.xlsx`);
const model = (records: XlsxRecords) => projectXlsxRecords(records).model;
describe("asynchronous native workbook record pipeline", () => {
  it("keeps Normal, explicit format zero and row/column defaults distinct in the independent style fixture", async () => {
    const records = await readXlsxRecords(readFileSync("packages/office/tests/fixtures/xlsx-cell-style-precedence.xlsx"));
    const book = model(records), sheet = book.snapshot().metadata.workbook.sheets[0]!;
    expect(sheet).toMatchObject({ defaultCssClass: "xlsx-named-0", defaultNumberFormatId: "num-164",
      rowDimensions: [{ from: 1, to: 1, cssClass: "xlsx-cell-2", numberFormatId: "num-167" }, { from: 4, to: 5, cssClass: "xlsx-cell-0", numberFormatId: "num-165" }],
      columnDimensions: [{ from: 1, to: 7, cssClass: "xlsx-cell-1", numberFormatId: "num-166" }] });
    for (let row = 0; row < 6; row++) {
      expect(book.getCell(sheet.sheetId, row, 1)?.cssClass).toBeUndefined();
      expect(book.getCell(sheet.sheetId, row, 2)).toMatchObject({ cssClass: "xlsx-cell-0", formatId: "num-165" });
      expect(book.getCell(sheet.sheetId, row, 4)?.value).toEqual({ type: "blank" });
      expect(book.getCell(sheet.sheetId, row, 5)).toBeUndefined();
      expect(book.getCell(sheet.sheetId, row, 6)).toMatchObject({ cssClass: "xlsx-cell-0", formatId: "num-165", value: { type: "blank" } });
    }
  });
  it("maps all 32 independent and producer cell streams while retaining unconsumed feature headers", async () => {
    const expected = JSON.parse(readFileSync("examples/xlsx/expected.json", "utf8")).cases as { file: string; cellCount: number; formulaCount: number }[];
    let checked = 0;
    for (const directory of ["examples/xlsx", "examples/xlsx/producer"]) for (const name of readdirSync(directory).filter(name => name.endsWith(".xlsx"))) {
      const records = await readXlsxRecords(readFileSync(`${directory}/${name}`));
      if (directory === "examples/xlsx") {
        const fixture = expected.find(item => item.file === name)!;
        expect(records.counts, name).toEqual({ cells: fixture.cellCount, formulas: fixture.formulaCount });
      }
      const book = model(records); expect(book.snapshot().metadata.workbook.sheets).toHaveLength(records.native.sheets.length);
      if (name.startsWith("15-")) {
        expect(records.counts).toEqual({ cells: 200000, formulas: 20000 }); expect(records.chunks.descriptors).toHaveLength(40);
        expect(book.getCell("sheet-1", 9999, 19)?.value).toEqual({ type: "number", value: 3_400_152 });
        expect(records.sheets[0]!.metadata.rowDimensions.length).toBeLessThan(3);
      }
      if (name.startsWith("05-") && directory === "examples/xlsx") expect(records.sheets[0]!.root.content).toContainEqual(expect.objectContaining({ name: "tableParts" }));
      checked++;
    }
    expect(checked).toBe(32);
  });
  it("checks options, global cell/formula/XML budgets and retained metadata limits before returning records", async () => {
    const bytes = source("04-formulas-names");
    for (const options of [
      { workbookLimits: { maxCells: 5 } }, { workbookLimits: { maxFormulas: 2 } },
      { nativeLimits: { maxXmlNodes: 50 } }, { nativeLimits: { maxRetainedNodes: 10 } },
      { nativeLimits: { maxMetadataBytes: 200 } }, { nativeLimits: { maxZipBytes: 100 } },
      { limits: { maxPackageBytes: 200 } },
    ]) await expect(readXlsxRecords(bytes, options)).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(readXlsxRecords(bytes, { nativeLimits: { maxXmlNodes: 0 } })).rejects.toMatchObject({ code: "InvalidOperation" });
    await expect(readXlsxRecords(bytes, { workbookLimits: { surprise: 1 } } as XlsxReadOptions)).rejects.toMatchObject({ code: "InvalidOperation" });
    // The 20 cells span two worksheets. A fresh per-sheet counter would allow it.
    await expect(readXlsxRecords(bytes, { workbookLimits: { maxCells: 19 } })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it("cancels at ZIP, metadata and worksheet phases and freezes host progress records", async () => {
    for (const phase of ["zip", "metadata", "worksheet"] as const) {
      const signal = { aborted: false }; let delivered = 0;
      await expect(readXlsxRecords(source("01-values-text"), { signal, onProgress(progress) {
        expect(Object.isFrozen(progress)).toBe(true); delivered++; if (progress.phase === phase) signal.aborted = true;
      } })).rejects.toMatchObject({ code: "OperationCancelled" }); expect(delivered).toBeGreaterThan(0);
    }
    await expect(readXlsxRecords(source("01-values-text"), { onProgress() { throw new Error("host failure"); } })).rejects.toMatchObject({ code: "HostSubmissionFailed" });
    const bytes = source("01-values-text"); let touched = false;
    const records = await readXlsxRecords(bytes, { onProgress() { if (!touched) { touched = true; bytes.fill(0); } } });
    expect(records.counts.cells).toBe(15); expect(model(records).getCell("sheet-1", 2, 0)?.value).toEqual({ type: "number", value: 12.5 });
  });
  it("fails a bad cell on the second pass without returning a partial candidate", async () => {
    const files = new Map(readDocumentZip(source("01-values-text"))), part = "xl/worksheets/sheet1.xml";
    files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace('t="s"><v>0</v>', 't="s"><v>99999</v>')));
    await expect(readXlsxRecords(writeOpcZip(files))).rejects.toMatchObject({ code: "InvalidDocument" });
  });
});
