import { describe, expect, it } from 'vitest';
import { workbookBuiltinTableStyle } from '../src/workbook/table-styles.js';

describe('normative predefined workbook table styles', () => {
  it('covers all 60 Annex G.1 names with immutable ordered regional differentials', () => {
    const precedence = ['wholeTable', 'firstColumnStripe', 'secondColumnStripe', 'firstRowStripe', 'secondRowStripe', 'lastColumn', 'firstColumn', 'headerRow', 'totalRow', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
    let count = 0;
    for (const [family, maximum] of [['Light', 21], ['Medium', 28], ['Dark', 11]] as const) for (let index = 1; index <= maximum; index++) {
      const name = `TableStyle${family}${index}`, style = workbookBuiltinTableStyle(name); count++;
      expect(style.name).toBe(name); expect(style).toBe(workbookBuiltinTableStyle(name));
      expect(Object.isFrozen(style)).toBe(true); expect(Object.isFrozen(style.elements)).toBe(true);
      expect(style.elements.length).toBeGreaterThan(0);
      expect(style.elements.map(element => precedence.indexOf(element.region))).toEqual(style.elements.map(element => precedence.indexOf(element.region)).sort((a, b) => a - b));
      for (const element of style.elements) {
        expect(element.size).toBe(1); expect(Object.isFrozen(element.style)).toBe(true);
        if (element.style.fill) expect(Object.isFrozen(element.style.fill)).toBe(true);
        if (element.style.borders) expect(Object.isFrozen(element.style.borders)).toBe(true);
      }
    }
    expect(count).toBe(60);
    for (const name of ['TableStyleLight0', 'TableStyleLight22', 'TableStyleMedium29', 'TableStyleDark12', 'PivotStyleLight1', '__proto__', 'constructor'])
      expect(() => workbookBuiltinTableStyle(name)).toThrow(/not supported/);
  });

  it('preserves Medium9 theme roles, exact tint values and regional inner borders', () => {
    const elements = workbookBuiltinTableStyle('TableStyleMedium9').elements;
    const region = (name: string) => elements.find(element => element.region === name)!.style;
    expect(region('wholeTable')).toEqual({ font: { color: { theme: 1, tint: 0 } }, fill: { theme: 4, tint: 0.79998168889431442 },
      borders: { vertical: { style: 'thin', color: { theme: 0, tint: 0 } }, horizontal: { style: 'thin', color: { theme: 0, tint: 0 } } } });
    expect(region('headerRow')).toEqual({ font: { bold: true, color: { theme: 0, tint: 0 } }, fill: { theme: 4, tint: 0 }, borders: { bottom: { style: 'thick', color: { theme: 0, tint: 0 } } } });
    expect(region('totalRow').borders).toEqual({ top: { style: 'thick', color: { theme: 0, tint: 0 } } });
    expect(region('firstRowStripe').fill).toEqual({ theme: 4, tint: 0.59999389629810485 });
    expect(region('firstColumnStripe')).toEqual(region('firstRowStripe'));
  });

  it('uses canonical Annex names for the pinned attachment errata without accepting those typos as source names', () => {
    expect(workbookBuiltinTableStyle('TableStyleLight14').elements).toHaveLength(9);
    expect(workbookBuiltinTableStyle('TableStyleDark9').elements).toHaveLength(7);
    for (const name of ['TableStyleLight14 ', 'TableStyleDark9 2']) expect(() => workbookBuiltinTableStyle(name)).toThrow(/not supported/);
  });
});
