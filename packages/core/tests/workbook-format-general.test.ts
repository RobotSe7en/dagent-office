import { describe, expect, it } from "vitest";
import { formatGeneralNumber } from "../src/workbook/format-general.js";

const show = (value: number, width = 100, decimal = '.') => formatGeneralNumber(value, decimal, width, text => text.length);
describe('bounded General number-format width fitting', () => {
  it('uses fixed notation through eleven digits and scientific notation at twelve, regardless of a wider column', () => {
    expect(show(12345678901)).toEqual({ kind: 'text', text: '12345678901', notation: 'decimal' });
    expect(show(123456789012)).toEqual({ kind: 'text', text: '1.23457E+11', notation: 'scientific' });
    expect(show(-12345678901)).toEqual({ kind: 'text', text: '-12345678901', notation: 'decimal' });
    expect(show(99999999999.5)).toEqual({ kind: 'text', text: '1E+11', notation: 'scientific' });
    expect(show(0)).toEqual({ kind: 'text', text: '0', notation: 'decimal' });
    expect(show(-0)).toEqual(show(0));
  });
  it('rounds fractional display to the available width without dropping integer places or turning small values into zero', () => {
    expect(show(25.76, 2)).toEqual({ kind: 'text', text: '26', notation: 'decimal' });
    expect(show(123.456, 6)).toEqual({ kind: 'text', text: '123.46', notation: 'decimal' });
    expect(show(123.456, 3)).toEqual({ kind: 'text', text: '123', notation: 'decimal' });
    expect(show(12345, 4)).toEqual({ kind: 'overflow' });
    expect(show(.00001, 6)).toEqual({ kind: 'text', text: '1E-05', notation: 'scientific' });
    expect(show(.00001, 4)).toEqual({ kind: 'overflow' });
    expect(show(.0000123456789)).toEqual({ kind: 'text', text: '1.23457E-05', notation: 'scientific' });
  });
  it('uses measured glyph widths and an explicit decimal symbol instead of treating the column as a character count', () => {
    expect(show(12.5, 4, ',')).toEqual({ kind: 'text', text: '12,5', notation: 'decimal' });
    expect(formatGeneralNumber(12.5, '.', 3.5, text => [...text].reduce((sum, character) => sum + (character === '.' ? .5 : 1), 0))).toEqual({ kind: 'text', text: '12.5', notation: 'decimal' });
    expect(show(-12.5, 4)).toEqual({ kind: 'text', text: '-13', notation: 'decimal' });
  });
  it('handles extreme finite numbers with a bounded number of measurements and rejects invalid metrics', () => {
    expect(show(Number.MIN_VALUE)).toEqual({ kind: 'text', text: '5E-324', notation: 'scientific' });
    expect(show(1e308)).toEqual({ kind: 'text', text: '1E+308', notation: 'scientific' });
    let calls = 0; expect(formatGeneralNumber(Math.PI, '.', 0, () => { calls++; return 1; })).toEqual({ kind: 'overflow' }); expect(calls).toBeLessThanOrEqual(25);
    expect(() => show(Infinity)).toThrow(); expect(() => show(1, -1)).toThrow();
    expect(() => formatGeneralNumber(1, '.', 1, () => NaN)).toThrow();
  });
});
