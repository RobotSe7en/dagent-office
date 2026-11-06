import { describe, expect, it } from "vitest";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { workbookViewport } from "../src/workbook/viewport.js";
import { hitWorkbookCell, projectWorkbookCells, projectionWindows, WorkbookMergeIndex } from "../src/workbook/projection.js";

const dimensions = { rows: new AxisDimensions(100, 10, []), columns: new AxisDimensions(100, 20, []) };
const range = (row: number, column: number, endRow: number, endColumn: number) => ({ from: { row, column }, to: { row: endRow, column: endColumn } });
describe("workbook source projection", () => {
  it("clips a merge across four panes without changing its source or text origin", () => {
    const merge = range(0, 0, 6, 6), index = new WorkbookMergeIndex([merge]);
    const layout = workbookViewport(dimensions, { rows: 2, columns: 1 }, { widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 }, 40);
    const cells = projectWorkbookCells(dimensions, layout, index, 40), parts = cells.filter(cell => cell.range === merge);
    expect(parts).toHaveLength(4); expect(parts.every(part => part.row === 0 && part.column === 0)).toBe(true);
    expect(parts.find(part => part.pane === 'body')).toMatchObject({ rect: { x: -40, y: -30, width: 140, height: 70 }, clip: { x: 20, y: 20, width: 80, height: 20 } });
    const windows = projectionWindows('sheet', cells);
    expect(windows).toEqual([{ sheetId: 'sheet', row: 0, column: 0, rows: 1, columns: 1 }, { sheetId: 'sheet', row: 7, column: 0, rows: 4, columns: 1 }, { sheetId: 'sheet', row: 7, column: 3, rows: 4, columns: 4 }]);
    expect(windows.reduce((sum, window) => sum + window.rows * window.columns, 0)).toBe(21);
    expect(hitWorkbookCell(dimensions, layout, index, 20, 20)).toEqual({ row: 0, column: 0 });
    expect(hitWorkbookCell(dimensions, layout, index, 20, 40)).toEqual({ row: 7, column: 3 });
  });
  it("reads the hidden offscreen anchor of a million-row merge only once", () => {
    const sparse = { rows: new AxisDimensions(1_048_576, 10, [{ from: 0, to: 1_048_574, sizePt: 0 }]), columns: new AxisDimensions(16_384, 20, []) };
    const index = new WorkbookMergeIndex([range(0, 0, 1_048_575, 1)]);
    const layout = workbookViewport(sparse, { rows: 0, columns: 0 }, { widthPt: 40, heightPt: 10, scrollLeftPt: 0, scrollTopPt: 0 }, 2);
    const cells = projectWorkbookCells(sparse, layout, index, 1);
    expect(cells).toHaveLength(1); expect(cells[0]?.rect).toEqual({ x: 0, y: 0, width: 40, height: 10 });
    expect(projectionWindows('sheet', cells)).toEqual([{ sheetId: 'sheet', row: 0, column: 0, rows: 1, columns: 1 }]);
    expect(hitWorkbookCell(sparse, layout, index, 35, 5)).toEqual({ row: 0, column: 0 });
  });
  it("finds nested row intervals in different columns and matches an independent source scan", () => {
    const ranges = [range(0, 0, 99, 0), range(10, 1, 11, 2), range(12, 1, 99, 2), ...Array.from({ length: 50 }, (_, i) => range(i * 2, 3, i * 2 + 1, 4))];
    const index = new WorkbookMergeIndex(ranges);
    for (let row = 0; row < 100; row++) for (let column = 0; column < 6; column++) {
      expect(index.at(row, column)).toBe(ranges.find(range => range.from.row <= row && range.to.row >= row && range.from.column <= column && range.to.column >= column));
    }
  });
  it("maps exact frozen edges and rejects excessive projections", () => {
    const layout = workbookViewport(dimensions, { rows: 2, columns: 1 }, { widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 }, 40), index = new WorkbookMergeIndex([]);
    expect(hitWorkbookCell(dimensions, layout, index, 19.99, 19.99)).toEqual({ row: 1, column: 0 });
    expect(hitWorkbookCell(dimensions, layout, index, 20, 20)).toEqual({ row: 5, column: 3 });
    expect(hitWorkbookCell(dimensions, layout, index, 100, 10)).toBeUndefined();
    expect(hitWorkbookCell(dimensions, layout, index, -1, 10)).toBeUndefined();
    expect(() => hitWorkbookCell(dimensions, layout, index, NaN, 0)).toThrow();
    expect(() => projectWorkbookCells(dimensions, layout, index, 39)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
  it("bounds point lookup for thousands of long merges in shuffled columns", () => {
    let probes = 0;
    const ranges = Array.from({ length: 10000 }, (_, index) => {
      const column = index * 7919 % 10000, source = range(index, column, 100000, column);
      return { get from() { probes++; return source.from; }, get to() { probes++; return source.to; } };
    });
    const index = new WorkbookMergeIndex(ranges); probes = 0;
    expect(index.at(50000, 7919)).toBe(ranges[1]);
    expect(probes).toBeLessThan(100);
  });
});
