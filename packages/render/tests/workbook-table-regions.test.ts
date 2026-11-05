import { describe, expect, it } from 'vitest';
import { WorkbookTableRegions, workbookTableRegionBorders, workbookTableRegionOrder } from '../src/workbook/table-regions.js';
import { workbookBuiltinTableStyle } from '../src/workbook/table-styles.js';
import type { WorkbookTableRegionElement } from '../src/workbook/table-regions.js';
import type { WorkbookTableRegion } from '../src/workbook/table-style-types.js';
import { AxisDimensions } from '../src/workbook/dimensions.js';

const table = { range: { from: { row: 4, column: 2 }, to: { row: 11, column: 6 } }, header: true, totals: true,
  rowStripes: true, columnStripes: true, showFirstColumn: true, showLastColumn: true };
const elements = [...workbookTableRegionOrder].reverse().map(region => ({ region, size: 1, style: region }));
const names = (regions: WorkbookTableRegions<string>, row: number, column: number) => regions.at({ row, column }).map(match => match.style);

describe('workbook table structural regions', () => {
  it('bands visible rows across hidden runs without resetting at the viewport or changing column stripes', () => {
    const rows = new AxisDimensions(1_048_576, 1, [{ from: 5, to: 6, sizePt: 0 }, { from: 9, to: 9, sizePt: 0 }]);
    const regions = new WorkbookTableRegions(table, [
      { region: 'firstRowStripe', size: 2, style: 'first' }, { region: 'secondRowStripe', size: 1, style: 'second' },
      { region: 'firstColumnStripe', size: 1, style: 'column' },
    ], rows);
    expect(names(regions, 7, 2)).toEqual(['column', 'first']); expect(names(regions, 8, 2)).toEqual(['column', 'first']);
    expect(regions.at({ row: 7, column: 2 }).at(-1)!.range).toEqual({ from: { row: 7, column: 2 }, to: { row: 8, column: 6 } });
    expect(names(regions, 10, 2)).toEqual(['column', 'second']); expect(names(regions, 9, 2)).toEqual(['column']);
    const almostHidden = new AxisDimensions(1_048_576, 0, [{ from: 500_000, to: 500_002, sizePt: 1 }]);
    const large = new WorkbookTableRegions({ ...table, header: false, totals: false, range: { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 1 } } }, elements, almostHidden);
    expect(names(large, 500_000, 0)).toContain('firstRowStripe'); expect(names(large, 500_001, 0)).toContain('secondRowStripe');
    expect(almostHidden.segmentCount).toBe(3);
  });
  it('applies the normative precedence at intersecting column, row and corner regions independently of source order', () => {
    const regions = new WorkbookTableRegions(table, elements);
    expect(names(regions, 4, 2)).toEqual(['wholeTable', 'firstColumnStripe', 'firstColumn', 'headerRow', 'firstHeaderCell']);
    expect(names(regions, 4, 6)).toEqual(['wholeTable', 'firstColumnStripe', 'lastColumn', 'headerRow', 'lastHeaderCell']);
    expect(names(regions, 5, 3)).toEqual(['wholeTable', 'secondColumnStripe', 'firstRowStripe']);
    expect(names(regions, 6, 4)).toEqual(['wholeTable', 'firstColumnStripe', 'secondRowStripe']);
    expect(names(regions, 11, 2)).toEqual(['wholeTable', 'firstColumnStripe', 'firstColumn', 'totalRow', 'firstTotalCell']);
    expect(names(regions, 11, 6)).toEqual(['wholeTable', 'firstColumnStripe', 'lastColumn', 'totalRow', 'lastTotalCell']);
    for (const [row, column] of [[3, 3], [12, 3], [6, 1], [6, 7]]) expect(names(regions, row!, column!)).toEqual([]);
  });

  it('keeps variable stripe sizes and clips the final band without enumerating the full worksheet', () => {
    const regions = new WorkbookTableRegions({ ...table, range: { from: { row: 0, column: 0 }, to: { row: 1_048_575, column: 16_383 } } }, [
      { region: 'firstRowStripe', size: 2, style: 'rows' }, { region: 'secondRowStripe', size: 1 },
      { region: 'firstColumnStripe', size: 9 }, { region: 'secondColumnStripe', size: 2, style: 'columns' },
    ]);
    expect(names(regions, 1, 9)).toEqual(['columns', 'rows']);
    expect(names(regions, 2, 10)).toEqual(['columns', 'rows']);
    expect(names(regions, 3, 11)).toEqual([]);
    expect(regions.at({ row: 1_048_573, column: 16_382 })).toEqual([
      { region: 'firstRowStripe', style: 'rows', range: { from: { row: 1_048_573, column: 0 }, to: { row: 1_048_574, column: 16_383 } } },
    ]);
    const clipped = new WorkbookTableRegions({ ...table, range: { from: { row: 4, column: 2 }, to: { row: 9, column: 6 } } }, [
      { region: 'firstRowStripe', size: 2, style: 'first' }, { region: 'secondRowStripe', size: 1, style: 'second' },
    ]);
    expect(clipped.at({ row: 8, column: 4 })[0]!.range).toEqual({ from: { row: 8, column: 2 }, to: { row: 8, column: 6 } });
  });

  it('handles headerless tables, absent totals, disabled bands and coincident first/last columns', () => {
    const plain = new WorkbookTableRegions({ ...table, header: false, totals: false, showFirstColumn: false, showLastColumn: false }, elements);
    expect(names(plain, 4, 2)).toEqual(['wholeTable', 'firstColumnStripe', 'firstRowStripe']);
    expect(names(plain, 11, 6)).toEqual(['wholeTable', 'firstColumnStripe', 'secondRowStripe']);
    const narrow = new WorkbookTableRegions({ ...table, range: { from: { row: 4, column: 2 }, to: { row: 5, column: 2 } }, rowStripes: false, columnStripes: false }, elements);
    expect(names(narrow, 4, 2)).toEqual(['wholeTable', 'lastColumn', 'firstColumn', 'headerRow', 'firstHeaderCell', 'lastHeaderCell']);
    expect(names(narrow, 5, 2)).toEqual(['wholeTable', 'lastColumn', 'firstColumn', 'totalRow', 'firstTotalCell', 'lastTotalCell']);
    const corners = new WorkbookTableRegions({ ...table, showFirstColumn: false, showLastColumn: false, rowStripes: false, columnStripes: false }, elements);
    expect(names(corners, 4, 2)).toEqual(['wholeTable', 'headerRow', 'firstHeaderCell']);
  });

  it('maps inner and outer regional borders without outlining every cell or losing explicit clearing', () => {
    const area = table.range, borders = { left: 'left', right: 'right', top: 'top', bottom: 'bottom', vertical: 'vertical', horizontal: 'horizontal' };
    expect(workbookTableRegionBorders(area, { row: 4, column: 2 }, borders)).toEqual({ before: 'left', after: 'vertical', top: 'top', bottom: 'horizontal' });
    expect(workbookTableRegionBorders(area, { row: 8, column: 4 }, borders)).toEqual({ before: 'vertical', after: 'vertical', top: 'horizontal', bottom: 'horizontal' });
    expect(workbookTableRegionBorders(area, { row: 11, column: 6 }, borders)).toEqual({ before: 'vertical', after: 'right', top: 'horizontal', bottom: 'bottom' });
    expect(workbookTableRegionBorders(area, { row: 8, column: 4 }, { left: 'none', horizontal: 'none' })).toEqual({ top: 'none', bottom: 'none' });
    const one = { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } };
    expect(workbookTableRegionBorders(one, one.from, borders)).toEqual({ before: 'left', after: 'right', top: 'top', bottom: 'bottom' });
    expect(() => workbookTableRegionBorders(area, { row: 3, column: 2 }, borders)).toThrow(/inside/);
  });

  it('composes Medium9 header and body fills/fonts and preserves logical internal borders', () => {
    const regions = new WorkbookTableRegions(table, workbookBuiltinTableStyle('TableStyleMedium9').elements);
    const resolve = (row: number, column: number) => {
      let font = {}, fill, borders = {};
      for (const match of regions.at({ row, column })) {
        font = { ...font, ...match.style.font }; fill = match.style.fill ?? fill;
        borders = { ...borders, ...workbookTableRegionBorders(match.range, { row, column }, match.style.borders ?? {}) };
      }
      return { font, fill, borders };
    };
    expect(resolve(4, 3)).toEqual({ font: { bold: true, color: { theme: 0, tint: 0 } }, fill: { theme: 4, tint: 0 }, borders: {
      before: { style: 'thin', color: { theme: 0, tint: 0 } }, after: { style: 'thin', color: { theme: 0, tint: 0 } }, bottom: { style: 'thick', color: { theme: 0, tint: 0 } },
    } });
    expect(resolve(5, 3).fill).toEqual({ theme: 4, tint: 0.59999389629810485 });
    expect(resolve(6, 3).fill).toEqual({ theme: 4, tint: 0.79998168889431442 });
  });

  it('snapshots structural inputs and rejects invalid coordinates and ambiguous definitions', () => {
    const source = structuredClone(table), entries = structuredClone(elements), regions = new WorkbookTableRegions(source, entries);
    source.range.from.row = 10; entries[0]!.size = 9;
    const result = regions.at({ row: 4, column: 2 }); expect(result.length).toBeGreaterThan(0);
    expect(Object.isFrozen(result)).toBe(true); expect(Object.isFrozen(result[0]!.range.from)).toBe(true);
    for (const row of [-1, 1_048_576, .5, NaN]) expect(() => regions.at({ row, column: 0 })).toThrow(/coordinate/);
    for (const bad of [0, 10, .5, NaN]) expect(() => new WorkbookTableRegions(table, [{ region: 'firstRowStripe', size: bad }])).toThrow(/regional/);
    expect(() => new WorkbookTableRegions(table, [elements[0]!, elements[0]!])).toThrow(/repeated/);
    expect(() => new WorkbookTableRegions(table, [{ region: 'unknown' as WorkbookTableRegion, size: 1 }])).toThrow(/regional/);
    expect(() => new WorkbookTableRegions({ ...table, range: { from: { row: 0, column: 0 }, to: { row: 0, column: 0 } } }, [])).toThrow(/geometry/);
    const unstyled: WorkbookTableRegionElement<string>[] = [{ region: 'firstRowStripe', size: 3 }, { region: 'secondRowStripe', size: 1, style: 'second' }];
    const gaps = new WorkbookTableRegions(table, unstyled);
    expect(names(gaps, 5, 3)).toEqual([]); expect(names(gaps, 7, 3)).toEqual([]); expect(names(gaps, 8, 3)).toEqual(['second']);
  });
});
