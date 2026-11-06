import { describe, expect, it } from 'vitest';
import { AxisDimensions } from '../src/workbook/dimensions.js';
import { workbookViewport } from '../src/workbook/viewport.js';
import { WorkbookMergeIndex, projectWorkbookCells } from '../src/workbook/projection.js';
import { projectWorkbookDiagonals } from '../src/workbook/diagonal-projection.js';

const paint = { kind: 'slantDashDot' as const, color: 'rgba(0,68,170,1)' };
describe('worksheet diagonal projection', () => {
  it('reads merged anchors once and keeps physical diagonal directions through four frozen panes', () => {
    const dimensions = { rows: new AxisDimensions(20, 30, []), columns: new AxisDimensions(20, 40, []) };
    const merges = new WorkbookMergeIndex([{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }]);
    const layout = workbookViewport(dimensions, { rows: 1, columns: 1 }, { widthPt: 120, heightPt: 90, scrollLeftPt: 20, scrollTopPt: 10 }, 100);
    const cells = projectWorkbookCells(dimensions, layout, merges, 100).filter(cell => !cell.row && !cell.column);
    expect(cells).toHaveLength(4);
    for (const rtl of [false, true]) {
      let reads = 0;
      const result = projectWorkbookDiagonals(cells, layout, (row, column) => { expect([row, column]).toEqual([0, 0]); reads++; return { up: paint, down: paint }; }, rtl);
      expect(reads).toBe(1); expect(Object.isFrozen(result)).toBe(true);
      let total = 0;
      for (const pane of result) for (const edge of pane.edges) {
        total++; expect(edge.to.x).toBeGreaterThan(edge.from.x);
        expect(Math.abs((edge.to.y - edge.from.y) / (edge.to.x - edge.from.x))).toBeCloseTo(.75);
        expect(Math.hypot(edge.to.x - edge.from.x, edge.to.y - edge.from.y)).toBeCloseTo(150);
        expect(edge.dashOffsetPt ?? 0).toBe(0);
      }
      expect(total).toBeGreaterThanOrEqual(4); expect(total).toBeLessThanOrEqual(8);
      expect(() => projectWorkbookDiagonals(cells, layout, () => ({ up: paint, down: paint }), rtl, 1)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    }
  });
  it('clips a sheet-sized merge to the viewport without expanding its interior', () => {
    const dimensions = { rows: new AxisDimensions(1_048_576, 12, []), columns: new AxisDimensions(16_384, 48, []) };
    const merges = new WorkbookMergeIndex([{ from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } }]);
    const layout = workbookViewport(dimensions, { rows: 0, columns: 0 }, { widthPt: 180, heightPt: 90, scrollLeftPt: 393_216, scrollTopPt: 6_291_456 }, 5000);
    const cells = projectWorkbookCells(dimensions, layout, merges, 5000); expect(cells).toHaveLength(1);
    let reads = 0; const result = projectWorkbookDiagonals(cells, layout, () => { reads++; return { down: paint }; }, false);
    expect(reads).toBe(1); expect(result).toHaveLength(1); expect(result[0]!.edges).toHaveLength(1);
    const edge = result[0]!.edges[0]!; expect(Math.hypot(edge.to.x - edge.from.x, edge.to.y - edge.from.y)).toBeLessThan(100);
    expect(edge.from.x).toBeCloseTo(-.046875); expect(edge.from.y).toBeCloseTo(-.75);
  });
});
