import { describe, expect, it } from "vitest";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { workbookViewport } from "../src/workbook/viewport.js";
import { projectWorkbookCells, projectionWindows, WorkbookMergeIndex } from "../src/workbook/projection.js";
import { workbookAppearanceCoordinates } from "../src/workbook/appearance-coordinates.js";

const dimensions = { rows: new AxisDimensions(1_048_576, 10, []), columns: new AxisDimensions(16_384, 20, []) };
describe("bounded worksheet appearance sources", () => {
  it("reads only corners and the visible perimeter axes of a sheet-sized merge", () => {
    const merge = { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } }, index = new WorkbookMergeIndex([merge]);
    const layout = workbookViewport(dimensions, { rows: 0, columns: 0 }, { widthPt: 60, heightPt: 40, scrollLeftPt: 1000, scrollTopPt: 1000 }, 12);
    const projections = projectWorkbookCells(dimensions, layout, index, 12), coordinates = workbookAppearanceCoordinates(dimensions, layout, projections, index, 18);
    expect(projections).toHaveLength(1); expect(coordinates).toHaveLength(18);
    expect(coordinates).toContainEqual({ row: 0, column: 0 }); expect(coordinates).toContainEqual({ row: 1_048_575, column: 16_383 });
    expect(coordinates).toContainEqual({ row: 103, column: 0 }); expect(coordinates).toContainEqual({ row: 0, column: 52 });
    expect(coordinates.some(cell => cell.row === 100 && cell.column === 50)).toBe(false);
    expect(projectionWindows('s', coordinates).reduce((sum, window) => sum + window.rows * window.columns, 0)).toBe(18);
    expect(Object.isFrozen(coordinates) && coordinates.every(Object.isFrozen)).toBe(true);
    expect(() => workbookAppearanceCoordinates(dimensions, layout, projections, index, 17)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
  it("includes the source corners of an adjacent merge even when that merge is outside the viewport", () => {
    const index = new WorkbookMergeIndex([{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }]);
    const layout = workbookViewport(dimensions, { rows: 0, columns: 0 }, { widthPt: 20, heightPt: 10, scrollLeftPt: 40, scrollTopPt: 30 }, 1);
    const projections = projectWorkbookCells(dimensions, layout, index, 1);
    expect(workbookAppearanceCoordinates(dimensions, layout, projections, index, 8)).toEqual([
      { row: 0, column: 0 }, { row: 0, column: 2 }, { row: 2, column: 0 }, { row: 2, column: 2 },
      { row: 3, column: 1 }, { row: 3, column: 2 }, { row: 3, column: 3 }, { row: 4, column: 2 },
    ]);
  });
  it("skips million-row hidden runs when finding adjacent displayed borders", () => {
    const rows = new AxisDimensions(1_048_576, 10, [{ from: 1, to: 1_048_574, sizePt: 0 }]), sparse = { ...dimensions, rows };
    expect(rows.neighbor(0, 1)).toBe(1_048_575); expect(rows.neighbor(1_048_575, -1)).toBe(0);
    expect(rows.neighbor(0, -1)).toBeUndefined(); expect(rows.neighbor(1_048_575, 1)).toBeUndefined();
    expect(new AxisDimensions(1_048_576, 0, []).neighbor(500_000, 1)).toBeUndefined();
    const index = new WorkbookMergeIndex([]), layout = workbookViewport(sparse, { rows: 0, columns: 0 }, { widthPt: 20, heightPt: 10, scrollLeftPt: 0, scrollTopPt: 10 }, 1);
    expect(workbookAppearanceCoordinates(sparse, layout, projectWorkbookCells(sparse, layout, index, 1), index, 3)).toEqual([
      { row: 0, column: 0 }, { row: 1_048_575, column: 0 }, { row: 1_048_575, column: 1 },
    ]);
    expect(rows.segmentCount).toBe(3);
  });
});
