import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { workbookCellShape, workbookCellShapeOffset } from '../src/workbook/cell-shape.js';
import { AxisDimensions } from '../src/workbook/dimensions.js';
import { workbookViewport } from '../src/workbook/viewport.js';
import { WorkbookMergeIndex } from '../src/workbook/projection.js';
import { projectWorkbookBorders } from '../src/workbook/border-projection.js';

describe('physical rotated worksheet cells', () => {
  it('matches independent native fill corners, including almost horizontal and diagonal-only cells', () => {
    const reference = JSON.parse(readFileSync('docs/evidence/xlsx-rotated-cells.json', 'utf8')) as { cases: { rotation: number; mode: string; cellSizePt: [number, number]; fillPointsPt: number[][] }[] };
    expect(reference.cases).toHaveLength(36);
    for (const item of reference.cases) {
      const rotation = item.rotation === 255 ? 0 : item.rotation <= 90 ? -item.rotation : item.rotation - 90;
      const shape = workbookCellShape(...item.cellSizePt, rotation, ['frame', 'top', 'left'].includes(item.mode));
      const points = [[shape.topOffsetPt, 0], [shape.widthPt + shape.topOffsetPt, 0], [shape.widthPt, shape.heightPt], [0, shape.heightPt]];
      for (let point = 0; point < 4; point++) for (let axis = 0; axis < 2; axis++) expect(Math.abs(points[point]![axis]! - item.fillPointsPt[point]![axis]!), JSON.stringify(item)).toBeLessThan(.12);
      expect(Object.isFrozen(shape)).toBe(true);
    }
  });
  it('keeps split merged borders on one slope and measures their dash phase along that slope in either column direction', () => {
    const dimensions = { rows: new AxisDimensions(10, 30, []), columns: new AxisDimensions(10, 40, []) };
    const range = { from: { row: 1, column: 1 }, to: { row: 3, column: 3 } }, merges = new WorkbookMergeIndex([range]);
    const layout = workbookViewport(dimensions, { rows: 0, columns: 0 }, { widthPt: 240, heightPt: 180, scrollLeftPt: 0, scrollTopPt: 0 }, 100);
    const border = { kind: 'dashed' as const, color: 'rgb(0,0,0)' };
    for (const rtl of [false, true]) {
      const edges = projectWorkbookBorders(dimensions, layout, merges, (r, c) => r === 1 && c === 1 ? { before: border, top: border } : {}, rtl, 100, () => -45)[0]!.edges;
      const sides = edges.filter(edge => edge.from.y !== edge.to.y).sort((a, b) => a.from.y - b.from.y);
      expect(sides).toHaveLength(3);
      const bottomX = rtl ? 200 : 40;
      for (const [index, edge] of sides.entries()) {
        expect(edge.from.x).toBeCloseTo(bottomX + 90 - index * 30); expect(edge.to.x).toBeCloseTo(bottomX + 60 - index * 30);
        expect(edge.dashOffsetPt).toBeCloseTo(index * Math.hypot(30, 30));
      }
      const tops = edges.filter(edge => edge.from.y === edge.to.y).sort((a, b) => a.from.x - b.from.x);
      expect(tops.map(edge => edge.dashOffsetPt)).toEqual([0, 40, 80]);
    }
  });
  it('rejects invalid angles and finite geometry overflow before producing a view', () => {
    for (const angle of [NaN, Infinity, 91, -91, .5]) expect(() => workbookCellShape(80, 60, angle, true)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    expect(() => workbookCellShape(80, Number.MAX_VALUE, 1, true)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(() => workbookCellShapeOffset(workbookCellShape(80, 60, -45, true), Infinity)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    const huge = workbookCellShape(1e12, 1e12, -1, true);
    expect(workbookCellShapeOffset(huge, huge.heightPt)).toBe(0);
  });
});
