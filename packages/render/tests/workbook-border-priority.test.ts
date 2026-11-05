import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { selectWorkbookBorder } from "../src/workbook/border-priority.js";
import type { WorkbookBorderKind } from "../src/workbook/border-priority.js";
import { workbookBorderStrokes } from "../src/workbook/border-style.js";

describe("regular worksheet shared-border selection", () => {
  it("matches 588 independent PDF edge colors, including RTL physical-side ties", () => {
    const report = JSON.parse(readFileSync("docs/evidence/xlsx-border-semantics.json", "utf8")) as { cases: { axis: "row" | "column"; rightToLeft: boolean; samples: { first: WorkbookBorderKind; second: WorkbookBorderKind; owners: string[] }[] }[] };
    let checked = 0;
    for (const test of report.cases) for (const sample of test.samples) {
      const first = Object.freeze({ kind: sample.first, color: "red" }), second = Object.freeze({ kind: sample.second, color: "green" });
      const result = selectWorkbookBorder(first, second, test.axis, test.rightToLeft);
      const owner = result === undefined ? [] : [result === first ? "first" : "second"];
      expect(owner, JSON.stringify({ axis: test.axis, rtl: test.rightToLeft, sample })).toEqual(sample.owners); checked++;
    }
    expect(checked).toBe(588);
  });
  it("preserves the chosen source and treats absent/none borders as nonpainting", () => {
    const visible = { kind: "double" as const }, none = { kind: "none" as const };
    expect(selectWorkbookBorder(visible, undefined, "row", false)).toBe(visible);
    expect(selectWorkbookBorder(none, visible, "column", true)).toBe(visible);
    expect(selectWorkbookBorder(undefined, none, "row", true)).toBeUndefined();
    expect(selectWorkbookBorder(none, none, "column", true)).toBeUndefined();
    expect(() => selectWorkbookBorder({ kind: "toString" as WorkbookBorderKind }, none, "row", false)).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
  it("compares source CSS widths with native outer widths without quantizing them", () => {
    const native = { kind: "double" as const }, medium = { kind: "medium" as const }, thin = { kind: "thin" as const };
    const css = { kind: "css" as const, style: "solid" as const, widthPt: 1.76 };
    expect(selectWorkbookBorder(native, css, "row", false)).toBe(css);
    expect(selectWorkbookBorder({ ...css, widthPt: 1.75 }, native, "row", false)).toBe(native);
    const double = { ...css, style: "double" as const, widthPt: .75 };
    expect(selectWorkbookBorder(thin, double, "column", true)).toBe(double);
    expect(selectWorkbookBorder(double, medium, "column", true)).toBe(medium);
    expect(selectWorkbookBorder(css, { ...css }, "column", true)).toBe(css);
    for (const widthPt of [0, -1, NaN, Infinity, 1e308]) expect(() => selectWorkbookBorder({ ...css, widthPt }, native, "row", false)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
  it("retains CSS stroke thickness, double gaps and round dotted geometry", () => {
    const solid = { kind: "css" as const, style: "solid" as const, widthPt: 2.25 };
    expect(workbookBorderStrokes(solid)).toEqual([{ width: 2.25, offset: 0 }]);
    const double = workbookBorderStrokes({ ...solid, style: "double" });
    expect(double.map(stroke => stroke.width)).toEqual([.75, .75]);
    expect(double[1]!.offset + double[1]!.width / 2 - (double[0]!.offset - double[0]!.width / 2)).toBe(2.25);
    expect(double[1]!.offset - double[1]!.width / 2 - (double[0]!.offset + double[0]!.width / 2)).toBe(.75);
    expect(workbookBorderStrokes({ ...solid, style: "dotted" })).toEqual([{ width: 2.25, offset: 0, dash: [0, 4.5], cap: "round" }]);
    expect(workbookBorderStrokes({ ...solid, style: "dashed" })).toEqual([{ width: 2.25, offset: 0, dash: [6.75, 2.25] }]);
    expect(() => workbookBorderStrokes({ ...solid, widthPt: 1e308 })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => workbookBorderStrokes({ kind: "invalid" as "slantDashDot" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
});
