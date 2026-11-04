import { describe, expect, it } from "vitest";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { WorkbookScrollAxis } from "../src/workbook/scroll.js";
import { projectWorkbookDrawing, workbookDrawingRect } from "../src/workbook/drawing.js";
import { workbookViewport } from "../src/workbook/viewport.js";

describe("workbook logical scroll and drawing geometry", () => {
  it("keeps normal scrolling at 96 dpi and reaches distant worksheet coordinates with a bounded track", () => {
    const small = new WorkbookScrollAxis(150000, 300);
    expect(small.toTrack(150)).toBeCloseTo(200, 8); expect(small.fromTrack(200)).toBeCloseTo(150, 8);
    const large = new WorkbookScrollAxis(1_048_576 * 409, 300);
    expect(large.trackSizePx).toBe(1_000_400); expect(large.maximumPt).toBe(428867284);
    for (const point of [0, 15, 150000, 300000000, large.maximumPt]) expect(large.fromTrack(large.toTrack(point))).toBeCloseTo(point, 6);
    expect(large.fromTrack(2_000_000)).toBe(large.maximumPt); expect(large.toTrack(-1)).toBe(0);
    expect(new WorkbookScrollAxis(0, 300).fromTrack(1)).toBe(0); expect(new WorkbookScrollAxis(0, 300).toTrack(1)).toBe(0);
    expect(() => large.fromTrack(NaN)).toThrow(); expect(() => new WorkbookScrollAxis(Infinity, 10)).toThrow();
  });
  it("resolves all three anchor modes with hidden rows and overflowing source offsets", () => {
    const dimensions = { rows: new AxisDimensions(100, 10, [{ from: 1, to: 8, sizePt: 0 }]), columns: new AxisDimensions(100, 20, []) };
    const from = { row: 9, column: 1, offsetXPt: -5, offsetYPt: 3 }, to = { row: 10, column: 3, offsetXPt: 7, offsetYPt: 20 };
    expect(workbookDrawingRect({ kind: 'absolute', xPt: -3, yPt: 20, widthPt: 50, heightPt: 40 }, dimensions)).toEqual({ x: -3, y: 20, width: 50, height: 40 });
    expect(workbookDrawingRect({ kind: 'one-cell', from, widthPt: 50, heightPt: 40 }, dimensions)).toEqual({ x: 15, y: 13, width: 50, height: 40 });
    expect(workbookDrawingRect({ kind: 'two-cell', from, to, moveWithCells: false, sizeWithCells: false }, dimensions)).toEqual({ x: 15, y: 13, width: 52, height: 27 });
    expect(() => workbookDrawingRect({ kind: 'two-cell', from: to, to: from, moveWithCells: true, sizeWithCells: true }, dimensions)).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
  it("clips the same source drawing in frozen and scrolled panes", () => {
    const dimensions = { rows: new AxisDimensions(100, 10, []), columns: new AxisDimensions(100, 20, []) };
    const layout = workbookViewport(dimensions, { rows: 2, columns: 1 }, { widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 }, 40);
    const parts = projectWorkbookDrawing({ x: 0, y: 0, width: 100, height: 100 }, 0, layout);
    expect(parts).toHaveLength(4);
    expect(parts[3]).toEqual({ pane: 'body', rect: { x: -40, y: -30, width: 100, height: 100 }, clip: { x: 20, y: 20, width: 40, height: 50 } });
  });
  it("does not cull a rotated object whose unrotated rectangle is outside the viewport", () => {
    const dimensions = { rows: new AxisDimensions(100, 10, []), columns: new AxisDimensions(100, 20, []) };
    const layout = workbookViewport(dimensions, { rows: 0, columns: 0 }, { widthPt: 100, heightPt: 80, scrollLeftPt: 0, scrollTopPt: 0 }, 40);
    const rect = { x: 110, y: 0, width: 10, height: 80 };
    expect(projectWorkbookDrawing(rect, 0, layout)).toEqual([]);
    expect(projectWorkbookDrawing(rect, 90, layout)[0]?.clip.x).toBeCloseTo(75);
    expect(projectWorkbookDrawing({ ...rect, width: 0 }, 90, layout)).toEqual([]);
    expect(() => projectWorkbookDrawing(rect, NaN, layout)).toThrow();
  });
});
