import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { tintWorkbookColor } from "../src/workbook/color.js";
import type { WorkbookRgb } from "../src/workbook/color.js";
const rgb = (hex: string): WorkbookRgb => [1, 3, 5].map(offset => parseInt(hex.slice(offset, offset + 2), 16)) as unknown as WorkbookRgb;

describe("worksheet HLS luminance colors", () => {
  it("matches independent continuous HLS and pinned native colors within the recorded 8-bit rounding bound", () => {
    const report = JSON.parse(readFileSync("docs/evidence/xlsx-color-tint-semantics.json", "utf8")) as { samples: { rgb: string; tint: number; native: string; continuousHls: string }[] };
    expect(report.samples).toHaveLength(108);
    for (const sample of report.samples) {
      const actual = tintWorkbookColor(rgb(sample.rgb), sample.tint);
      for (const expected of [rgb(sample.native), rgb(sample.continuousHls)]) {
        expect(Math.max(...actual.map((channel, index) => Math.abs(channel - expected[index]!))), JSON.stringify(sample)).toBeLessThanOrEqual(1);
      }
      if (sample.tint === 0) expect(actual).toEqual(rgb(sample.rgb));
    }
  });
  it("keeps endpoints and hue/saturation distinct from RGB white mixing", () => {
    expect(tintWorkbookColor([255, 0, 0], -.5)).toEqual([128, 0, 0]);
    expect(tintWorkbookColor([68, 114, 196], .4)).toEqual([143, 170, 220]);
    expect(tintWorkbookColor([68, 114, 196], -.2)).toEqual([51, 90, 161]);
    expect(tintWorkbookColor([68, 114, 196], -1)).toEqual([0, 0, 0]);
    expect(tintWorkbookColor([68, 114, 196], 1)).toEqual([255, 255, 255]);
    expect(Object.isFrozen(tintWorkbookColor([1, 2, 3], 0))).toBe(true);
  });
  it("rejects malformed channels and tint instead of silently substituting black", () => {
    for (const tint of [-1.1, 1.1, NaN, Infinity]) expect(() => tintWorkbookColor([0, 0, 0], tint)).toThrow();
    for (const channels of [[-1, 2, 3], [0, 1, 256], [0, .5, 2], [0, NaN, 1]]) expect(() => tintWorkbookColor(channels as unknown as WorkbookRgb, 0)).toThrow();
  });
});
