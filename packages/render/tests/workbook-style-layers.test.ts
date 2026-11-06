import { describe, expect, it } from 'vitest';
import { workbookStyleDifferences } from '../src/workbook/style-layers.js';

describe('workbook appearance differences', () => {
  it('keeps changed paints together without treating reserved native alpha as display opacity', () => {
    const normal = new Map([['color', 'black'], ['--office-cell-foreground-color', '#000000'], ['--office-cell-foreground-tint', '0'],
      ['--office-cell-foreground-native-alpha', '0'], ['font-weight', '400']]);
    const alphaOnly = new Map(normal); alphaOnly.set('--office-cell-foreground-native-alpha', '255');
    expect(workbookStyleDifferences(alphaOnly, normal).size).toBe(0);
    const red = new Map(alphaOnly); red.set('--office-cell-foreground-color', '#ff0000'); red.set('color', 'red');
    const result = workbookStyleDifferences(red, normal);
    expect([...result]).toEqual([...red].filter(([name]) => name !== 'font-weight'));
  });
  it('restores a fill and logical border side without replacing unchanged font properties', () => {
    const normal = new Map([['--office-cell-fill-kind', 'none'], ['--office-cell-fill-tint', '0'], ['--office-cell-border-left-kind', 'none'],
      ['--office-cell-border-start-kind', ''], ['--office-cell-border-left-color', '#000000'], ['font-weight', '400']]);
    const source = new Map(normal); source.set('--office-cell-fill-kind', 'solid'); source.set('--office-cell-fill-color', '#123456'); source.set('--office-cell-border-start-kind', 'thin');
    expect([...workbookStyleDifferences(source, normal)]).toEqual([...source].filter(([name]) => name !== 'font-weight'));
  });
});
