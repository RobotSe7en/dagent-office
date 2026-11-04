import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { selectWorkbookBorder } from "../src/workbook/border-priority.js";
import type { WorkbookBorderKind } from "../src/workbook/border-priority.js";

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
});
