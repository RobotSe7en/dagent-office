import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { AxisDimensions } from '../src/workbook/dimensions.js';
import { workbookViewport } from '../src/workbook/viewport.js';
import { WorkbookMergeIndex, projectWorkbookCells } from '../src/workbook/projection.js';
import { workbookAppearanceCoordinates } from '../src/workbook/appearance-coordinates.js';
import { projectWorkbookBorders } from '../src/workbook/border-projection.js';
import type { WorkbookCellBorders } from '../src/workbook/cell-borders.js';

const red = { kind: 'thin' as const, color: 'rgba(255,0,0,1)' }, green = { kind: 'medium' as const, color: 'rgba(0,255,0,1)' };
const dimensions = { rows: new AxisDimensions(10, 30, []), columns: new AxisDimensions(10, 40, []) };
const frozen = { rows: 0, columns: 0 }, viewport = { widthPt: 200, heightPt: 150, scrollLeftPt: 0, scrollTopPt: 0 };
const side = { left: 'before', right: 'after', top: 'top', bottom: 'bottom' } as const;

describe('bounded worksheet border composition', () => {
  it('matches all isolated merged-side presence cases from the independent native PDF', () => {
    const report = JSON.parse(readFileSync('docs/evidence/xlsx-merged-border-semantics.json', 'utf8')) as { cases: { rightToLeft: boolean; sourceCell?: string; sourceSide: keyof typeof side; pdfStrokes: unknown[] }[] };
    const merges = new WorkbookMergeIndex([{ from: { row: 0, column: 0 }, to: { row: 2, column: 2 } }]), layout = workbookViewport(dimensions, frozen, viewport, 100);
    let checked = 0;
    for (const item of report.cases) {
      if (!item.sourceCell) continue;
      const column = item.sourceCell.charCodeAt(0) - 65, row = Number(item.sourceCell.slice(1)) - 1;
      const result = projectWorkbookBorders(dimensions, layout, merges, (r, c) => r === row && c === column ? { [side[item.sourceSide]]: red } : {}, item.rightToLeft)[0]!.edges;
      // Three constituent source columns/rows cover one full merge boundary.
      expect(result.length > 0, JSON.stringify(item)).toBe(item.pdfStrokes.length > 0);
      if (result.length) {
        expect(result).toHaveLength(3);
        expect(result.every(edge => edge.color === red.color)).toBe(true);
        const vertical = item.sourceSide === 'left' || item.sourceSide === 'right';
        expect(result.reduce((sum, edge) => sum + (vertical ? edge.to.y - edge.from.y : edge.to.x - edge.from.x), 0)).toBe(vertical ? 90 : 120);
      }
      checked++;
    }
    expect(checked).toBe(72);
  });
  it('lets a stronger outside neighbor replace only its segment and retains merged dash phase', () => {
    const range = { from: { row: 1, column: 1 }, to: { row: 3, column: 3 } }, merges = new WorkbookMergeIndex([range]);
    const layout = workbookViewport(dimensions, frozen, { ...viewport, widthPt: 240, heightPt: 180 }, 100);
    const read = (row: number, column: number): WorkbookCellBorders => {
      const out: { top?: typeof green | { kind: 'dashed'; color: string }; bottom?: typeof green; before?: typeof red; after?: typeof red } = {};
      if (row === 1 && column === 1) { out.top = { ...red, kind: 'dashed' }; out.before = red; }
      if (row === 1 && column === 3) out.after = red;
      if (row === 3 && column === 1) out.bottom = { ...green, kind: 'medium' };
      if (row === 0 && column === 2) out.bottom = green;
      return out;
    };
    for (const rtl of [false, true]) {
      const edges = projectWorkbookBorders(dimensions, layout, merges, read, rtl)[0]!.edges;
      const top = edges.filter(edge => edge.from.y === 30 && edge.to.y === 30).sort((a, b) => a.from.x - b.from.x);
      expect(top.map(edge => edge.kind)).toEqual(['dashed', 'medium', 'dashed']);
      expect(top.map(edge => edge.to.x - edge.from.x)).toEqual([40, 40, 40]);
      expect([top[0]!.dashOffsetPt, top[2]!.dashOffsetPt]).toEqual([0, 80]);
      expect(edges).toHaveLength(12);
    }
  });
  it('retains all eight native stronger-neighbor cases on their physical merge sides', () => {
    const report = JSON.parse(readFileSync('docs/evidence/xlsx-merged-border-semantics.json', 'utf8')) as { cases: { rightToLeft: boolean; outsideCell?: string; outsideSide: keyof typeof side; mergedSide: keyof typeof side }[] };
    const range = { from: { row: 1, column: 1 }, to: { row: 3, column: 3 } }, merges = new WorkbookMergeIndex([range]);
    const layout = workbookViewport(dimensions, frozen, { ...viewport, widthPt: 240, heightPt: 180 }, 100); let checked = 0;
    for (const item of report.cases) {
      if (!item.outsideCell) continue;
      const column = item.outsideCell.charCodeAt(0) - 65, row = Number(item.outsideCell.slice(1)) - 1;
      const read = (r: number, c: number): WorkbookCellBorders => r === row && c === column ? { [side[item.outsideSide]]: green }
        : r >= 1 && r <= 3 && c >= 1 && c <= 3 ? { top: red, bottom: red, before: red, after: red } : {};
      const edges = projectWorkbookBorders(dimensions, layout, merges, read, item.rightToLeft)[0]!.edges;
      expect(edges).toHaveLength(12); const incoming = edges.filter(edge => edge.color === green.color); expect(incoming).toHaveLength(1);
      const edge = incoming[0]!;
      if (item.mergedSide === 'top' || item.mergedSide === 'bottom') {
        expect(edge.from.y).toBe(item.mergedSide === 'top' ? 30 : 120); expect(edge.to.x - edge.from.x).toBe(40);
      } else {
        const x = item.mergedSide === 'left' ? 40 : 160; expect(edge.from.x).toBe(item.rightToLeft ? 240 - x : x); expect(edge.to.y - edge.from.y).toBe(30);
      }
      checked++;
    }
    expect(checked).toBe(8);
  });
  it('uses only bounded appearance sources for frozen fragments and a sheet-sized merge', () => {
    const large = { rows: new AxisDimensions(1_048_576, 12, []), columns: new AxisDimensions(16_384, 48, []) };
    const merges = new WorkbookMergeIndex([{ from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } }]);
    const layout = workbookViewport(large, { rows: 2, columns: 1 }, { widthPt: 180, heightPt: 90, scrollLeftPt: 2000, scrollTopPt: 120_000 }, 5000);
    const projections = projectWorkbookCells(large, layout, merges, 5000), sources = workbookAppearanceCoordinates(large, layout, projections, merges, 1000);
    const permitted = new Set(sources.map(point => point.row * 16_384 + point.column)), reads = new Set<number>();
    const result = projectWorkbookBorders(large, layout, merges, (row, column) => {
      const key = row * 16_384 + column; expect(permitted.has(key)).toBe(true); reads.add(key); return { top: red, before: red, after: red, bottom: red };
    }, false);
    expect(reads.size).toBeLessThanOrEqual(4); expect(result).toHaveLength(4);
    expect(result.find(pane => pane.pane === 'body')!.edges).toHaveLength(0);
    expect(result.find(pane => pane.pane === 'corner')!.edges.length).toBeGreaterThan(0);
  });
  it('skips hidden runs, applies physical RTL ties and detects missing sources/budget failures', () => {
    const sparse = { rows: new AxisDimensions(1_048_576, .1, [{ from: 1, to: 1_048_574, sizePt: 0 }]), columns: new AxisDimensions(3, .1, []) };
    const layout = workbookViewport(sparse, frozen, { widthPt: 1, heightPt: 1, scrollLeftPt: 0, scrollTopPt: 0 }, 100), merges = new WorkbookMergeIndex([]);
    const read = () => ({ before: { ...red, color: 'second' }, after: { ...red, color: 'first' }, top: red, bottom: red });
    const ltr = projectWorkbookBorders(sparse, layout, merges, read, false)[0]!, rtl = projectWorkbookBorders(sparse, layout, merges, read, true)[0]!;
    expect(ltr.edges).toHaveLength(17); expect(rtl.edges).toHaveLength(17);
    expect(ltr.edges.filter(edge => edge.from.x === .1 && edge.to.x === .1).every(edge => edge.color === 'second')).toBe(true);
    expect(rtl.edges.filter(edge => Math.abs(edge.from.x - .9) < 1e-10 && edge.to.x === edge.from.x).every(edge => edge.color === 'first')).toBe(true);
    expect(() => projectWorkbookBorders(sparse, layout, merges, () => undefined, false)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
    expect(() => projectWorkbookBorders(sparse, layout, merges, read, false, 16)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
    expect(() => projectWorkbookBorders(sparse, layout, merges, read, false, 0)).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
  });
});
