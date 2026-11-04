import { describe, expect, it } from "vitest";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { formatDecimalSection } from "../src/workbook/format-numeric.js";
import type { NumberSeparators } from "../src/workbook/format-numeric.js";

const separators = { decimal: '.', group: ',', primaryGroupSize: 3, secondaryGroupSize: 3 };
function display(value: number, code: string, symbols: NumberSeparators = separators) { return formatDecimalSection(value, compileNumberFormat(code)[0]!, symbols); }
function text(value: number, code: string): string { return display(value, code).parts.map(part => part.kind === 'text' ? part.text : part.kind === 'space' && part.character === '0' ? '\u2007' : '<layout>').join(''); }
describe('decimal number-format display parts', () => {
  it.each([
    [1.005, '0.00', '1.01'], [12345.678, '#,##0.00', '12,345.68'], [1234.5, '"¥"#,##0.00', '¥1,234.50'],
    [12200000, '0.0,,"M"', '12.2M'], [.5, '0.00%%', '50.00%%'], [.47, '#.##', '.47'],
    [0, '#.##', ''], [0, '0.##', '0'], [0, '#.??', '.\u2007\u2007'], [1, '0#', '01'], [1, '?0#', '\u200701'],
    [1234, '#,##00000', '01,234'], [1, '0.', '1'], [1.1, '0.#0', '1.10'], [1.12, '0.0?', '1.12'], [1.1, '0.0?', '1.1\u2007'],
    [0, '0#', '0'], [0, '0?', '0\u2007'], [0, '?0#', '\u20070'], [0, '00?', '00\u2007'], [0, '?#0', '\u20070'],
    [0, '0.?0', '0.\u20070'], [1, '0.?0', '1.\u20070'], [0, '.00', '.00'], [0, '0.#0', '0.00'], [1, '0.#0', '1.00'],
  ] as const)('formats %s with %s', (value, code, expected) => { expect(text(value, code)).toBe(expected); });
  it('retains measured spacing and fill instead of guessing a character count', () => {
    expect(display(12, '_)* 0').parts).toEqual([{ kind: 'space', character: ')' }, { kind: 'fill', character: ' ' }, { kind: 'text', text: '12' }]);
    expect(formatDecimalSection(-1234.5, compileNumberFormat('0;[Red](#,##0.00)')[1]!, separators).parts).toEqual([{ kind: 'text', text: '(1,234.50)' }]);
  });
  it('uses explicit grouping and decimal symbols without changing the source format code', () => {
    expect(display(12345.678, '#,##0.00', { ...separators, decimal: ',', group: '.' }).parts).toEqual([{ kind: 'text', text: '12.345,68' }]);
    expect(display(12345678, '#,##0', { ...separators, secondaryGroupSize: 2 }).parts).toEqual([{ kind: 'text', text: '1,23,45,678' }]);
  });
  it('reports rounding to zero separately from the sign and selected section', () => {
    expect(display(-.001, '0.00')).toEqual({ parts: [{ kind: 'text', text: '0.00' }], zero: true });
    expect(display(-.01, '0.00').zero).toBe(false);
    expect(text(1e308, '0%')).toBe('1' + '0'.repeat(310) + '%');
  });
  it('rejects other pattern kinds instead of silently applying a numeric fallback', () => {
    for (const code of ['yyyy-mm-dd', '0.00E+00', '# ?/?', 'General', '000-00-0000', '0,.00', '0..0']) expect(() => display(1, code), code).toThrowError(expect.objectContaining({ code: 'UnsupportedFeature' }));
    expect(() => display(1, '0', { ...separators, secondaryGroupSize: 0 })).toThrow();
  });
});
