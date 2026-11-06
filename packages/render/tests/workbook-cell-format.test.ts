import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookCell } from "@dagent-office/core";
import { workbookCellFormat } from "../src/workbook/cell-format.js";

const base = openDocument(createWorkbook({ documentId: "cell-formats" }), { kind: "workbook" }).snapshot().metadata.workbook.sheets[0]!;
const cell = (row: number, column: number, props: Partial<WorkbookCell> = {}): WorkbookCell => ({ sheetId: base.sheetId, row, column, value: { type: "blank" }, html: "", text: "", ...props });

describe("worksheet source format references", () => {
  it("matches independent native precedence for stored and implicit blanks, including explicit format zero", () => {
    const reference = JSON.parse(readFileSync("docs/evidence/xlsx-cell-style-semantics.json", "utf8")) as { samples: { address: string; text: string; fill: string }[] };
    const sheet = { ...base, defaultCssClass: "normal", defaultNumberFormatId: "N", columnDimensions: [{ from: 1, to: 7, unit: "characters" as const, cssClass: "column", numberFormatId: "C" }],
      rowDimensions: [{ from: 1, to: 1, cssClass: "row", numberFormatId: "R" }, { from: 4, to: 5, cssClass: "zero", numberFormatId: "D" }] };
    const fills: Record<string, string> = { normal: "#eeeeee", zero: "#ffffdd", column: "#ddeeff", row: "#ffdddd", explicit: "#ddffdd" };
    expect(reference.samples).toHaveLength(54);
    for (const sample of reference.samples) {
      const row = Number(sample.address.slice(1)) - 1, column = sample.address.charCodeAt(0) - 65;
      const source = column === 5 ? undefined : cell(row, column, [2, 6].includes(column) ? { cssClass: "zero", formatId: "D" } : [3, 7].includes(column) ? { cssClass: "explicit", formatId: "X" } : {});
      const format = workbookCellFormat(sheet, { row, column }, source);
      expect(fills[format.cssClass], sample.address).toBe(sample.fill);
      if (sample.text) expect(format.numberFormatId + "1", sample.address).toBe(sample.text);
    }
  });
  it("resolves appearance and number references independently without concatenating CSS classes", () => {
    const sheet = { ...base, defaultCssClass: "normal", defaultNumberFormatId: "N",
      columnDimensions: [{ from: 0, to: 16_383, unit: "characters" as const, cssClass: "column", numberFormatId: "C" }],
      rowDimensions: [{ from: 0, to: 1_048_574, hidden: true, cssClass: "row" }, { from: 1_048_575, to: 1_048_575, heightPt: 22, numberFormatId: "last" }] };
    expect(workbookCellFormat(sheet, { row: 100, column: 0 })).toEqual({ cssClass: "row", numberFormatId: "C" });
    expect(workbookCellFormat(sheet, { row: 1_048_575, column: 16_383 })).toEqual({ cssClass: "column", numberFormatId: "last" });
    expect(workbookCellFormat(sheet, { row: 1, column: 1 }, cell(1, 1, { cssClass: "one two", formatId: "own" }))).toEqual({ cssClass: "one two", numberFormatId: "own" });
    expect(Object.isFrozen(workbookCellFormat(base, { row: 0, column: 0 }))).toBe(true);
    expect(workbookCellFormat({ ...base, defaultNumberFormatId: undefined }, { row: 0, column: 0 })).not.toHaveProperty("numberFormatId");
  });
  it("rejects foreign cells and invalid coordinates rather than styling a different source", () => {
    for (const coordinate of [{ row: -1, column: 0 }, { row: 0, column: 16_384 }, { row: 1.5, column: 0 }, { row: Infinity, column: 0 }]) expect(() => workbookCellFormat(base, coordinate)).toThrow();
    expect(() => workbookCellFormat(base, { row: 0, column: 0 }, cell(1, 0))).toThrow();
    expect(() => workbookCellFormat(base, { row: 0, column: 0 }, cell(0, 0, { sheetId: "other" }))).toThrow();
  });
});
