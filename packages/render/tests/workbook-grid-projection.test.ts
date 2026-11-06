import { describe, expect, it } from 'vitest';
import { AxisDimensions } from '../src/workbook/dimensions.js';
import { workbookViewport } from '../src/workbook/viewport.js';
import { projectWorkbookGridLines, WORKBOOK_GRID_PAINT } from '../src/workbook/grid-projection.js';
import { projectWorkbookBorders } from '../src/workbook/border-projection.js';
import { WorkbookMergeIndex } from '../src/workbook/projection.js';

describe('uniform worksheet grid projection', () => {
  it('preserves the complete segmented grid coverage through hidden dimensions, partial cells and frozen panes', () => {
    const dimensions = { rows: new AxisDimensions(100, 17.3, [{ from: 2, to: 4, sizePt: 0 }, { from: 8, to: 8, sizePt: 29.1 }]),
      columns: new AxisDimensions(50, 23.7, [{ from: 1, to: 1, sizePt: 0 }, { from: 5, to: 6, sizePt: 41.2 }]) };
    const side = WORKBOOK_GRID_PAINT, read = () => ({ before: side, after: side, top: side, bottom: side });
    for (const rightToLeft of [false, true]) for (const frozen of [{ rows: 0, columns: 0 }, { rows: 2, columns: 3 }]) for (const offset of [0, 9.1, 102.7, 10000]) {
      const layout = workbookViewport(dimensions, frozen, { widthPt: 187.4, heightPt: 113.6, scrollLeftPt: offset, scrollTopPt: offset }, 5000);
      const before = JSON.stringify(layout), expected = projectWorkbookBorders(dimensions, layout, new WorkbookMergeIndex([]), read, rightToLeft);
      const actual = projectWorkbookGridLines(layout, rightToLeft);
      for (const [index, pane] of actual.entries()) {
        expect(pane.bounds).toEqual(expected[index]!.bounds); expect(pane.pane).toBe(expected[index]!.pane);
        for (const edge of expected[index]!.edges) expect(pane.edges.some(full => {
          const horizontal = edge.from.y === edge.to.y;
          return horizontal ? full.from.y === edge.from.y && full.to.y === edge.to.y && full.from.x <= edge.from.x + 1e-9 && full.to.x >= edge.to.x - 1e-9
            : full.from.x === edge.from.x && full.to.x === edge.to.x && full.from.y <= edge.from.y + 1e-9 && full.to.y >= edge.to.y - 1e-9;
        })).toBe(true);
        expect(pane.edges.length).toBeLessThanOrEqual(layout.panes[index]!.rows.length + layout.panes[index]!.columns.length + 2);
      }
      expect(JSON.stringify(layout)).toBe(before);
    }
  });
  it('bounds the combined physical edges across every pane', () => {
    const dimensions = { rows: new AxisDimensions(100, 10, []), columns: new AxisDimensions(100, 20, []) };
    const layout = workbookViewport(dimensions, { rows: 1, columns: 1 }, { widthPt: 100, heightPt: 100, scrollLeftPt: 5, scrollTopPt: 5 }, 5000);
    const count = projectWorkbookGridLines(layout, false).reduce((sum, pane) => sum + pane.edges.length, 0);
    expect(projectWorkbookGridLines(layout, false, count)).toHaveLength(4);
    expect(() => projectWorkbookGridLines(layout, false, count - 1)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(() => projectWorkbookGridLines(layout, false, 0)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
  });
});
