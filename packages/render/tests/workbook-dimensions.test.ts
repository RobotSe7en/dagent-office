import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import { AxisDimensions, columnWidthPt, worksheetDimensions } from "../src/workbook/dimensions.js";

describe("sparse worksheet dimensions", () => {
  it("matches the native 96 dpi width example without inventing additional padding", () => {
    expect(columnWidthPt(8.7109375, 7)).toBe(61 * 0.75);
    expect(columnWidthPt(8, 7, "baseCharacters")).toBe(61 * 0.75);
    expect(columnWidthPt(8, 7)).toBe(56 * 0.75); expect(columnWidthPt(0, 7)).toBe(0);
  });
  it("skips huge hidden spans while retaining distant row identities and exact edges", () => {
    const axis = new AxisDimensions(1_048_576, 15, [{ from: 1, to: 1_048_574, sizePt: 0 }]);
    expect(axis.segmentCount).toBe(3); expect(axis.extentPt).toBe(30);
    expect(axis.hit(14.99)).toBe(0); expect(axis.hit(15)).toBe(1_048_575); expect(axis.hit(30)).toBeUndefined();
    expect(axis.position(1)).toBe(15); expect(axis.position(1_048_575)).toBe(15); expect(axis.position(1_048_576)).toBe(30);
    expect(axis.visible(-5, 30, 2)).toEqual([{ index: 0, startPt: 0, sizePt: 15 }, { index: 1_048_575, startPt: 15, sizePt: 15 }]);
    expect(axis.visible(15, 15, 1)).toEqual([]); expect(axis.visible(30, 40, 1)).toEqual([]);
  });
  it("agrees with an independent expanded reference across fractional, hidden and zero-size overrides", () => {
    const runs = [{ from: 0, to: 2, sizePt: 0 }, { from: 8, to: 15, sizePt: 13.3 }, { from: 16, to: 18, sizePt: 0 }, { from: 25, to: 26, sizePt: 20.75 }];
    const axis = new AxisDimensions(40, 15, runs), sizes = Array.from({ length: 40 }, (_, index) => runs.find(run => index >= run.from && index <= run.to)?.sizePt ?? 15);
    let position = 0;
    for (const [index, size] of sizes.entries()) {
      expect(axis.size(index)).toBe(size); expect(axis.position(index)).toBeCloseTo(position, 8);
      if (size) { expect(axis.hit(axis.position(index))).toBe(index); expect(axis.hit(axis.position(index) + size / 2)).toBe(index); }
      position += size;
    }
    const visible = axis.visible(45, 220, 40).map(cell => cell.index);
    let sum = 0; const expected: number[] = [];
    for (const [index, size] of sizes.entries()) { if (size && sum < 220 && sum + size > 45) expected.push(index); sum += size; }
    expect(visible).toEqual(expected);
  });
  it("uses the source sheet defaults and explicit unhidden exceptions", () => {
    const model = openDocument(createWorkbook({ documentId: "dimensions" }), { kind: "workbook" }), source = model.snapshot().metadata.workbook.sheets[0]!;
    const dimensions = worksheetDimensions({ ...source, defaultRowsHidden: true, rowDimensions: [{ from: 200, to: 200, hidden: false, heightPt: 18 }],
      defaultColumnWidth: 8, defaultColumnWidthUnit: "baseCharacters", columnDimensions: [{ from: 0, to: 2, unit: "characters", hidden: true }] }, 7);
    expect(dimensions.rows.extentPt).toBe(18); expect(dimensions.rows.hit(0)).toBe(200); expect(dimensions.rows.visible(0, 18, 1)[0]?.index).toBe(200);
    expect(dimensions.columns.hit(0)).toBe(3); expect(dimensions.columns.size(3)).toBe(45.75);
    expect(new AxisDimensions(100, 0, []).visible(0, 100, 1)).toEqual([]);
  });
  it("rejects invalid geometry and bounds dense viewports before expanding the sheet", () => {
    expect(() => new AxisDimensions(10, 15, [{ from: 1, to: 5, sizePt: 10 }, { from: 5, to: 8, sizePt: 20 }])).toThrow();
    expect(() => new AxisDimensions(10, NaN, [])).toThrow(); expect(() => columnWidthPt(8, 0)).toThrow();
    const axis = new AxisDimensions(1_048_576, 15, []);
    expect(axis.segmentCount).toBe(1); expect(() => axis.visible(0, axis.extentPt, 5000)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => axis.position(-1)).toThrow(); expect(() => axis.hit(Infinity)).toThrow(); expect(axis.hit(-1)).toBeUndefined();
  });
});
