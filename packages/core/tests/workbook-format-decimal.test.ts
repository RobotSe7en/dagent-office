import { describe, expect, it } from "vitest";
import { decimalMagnitude } from "../src/workbook/format-decimal.js";

describe('bounded decimal display rounding', () => {
  it.each([
    [1.005, 2, 0, '1', '01'], [-1.005, 2, 0, '1', '01'], [9.999, 2, 0, '10', '00'],
    [12345.678, 2, 0, '12345', '68'], [0.47, 2, 0, '0', '47'], [0, 2, 0, '0', '00'],
    [0.125, 2, 2, '12', '50'], [12200000, 1, -6, '12', '2'], [1234, 2, -3, '1', '23'],
    [1.23456e-7, 8, 0, '0', '00000012'], [0.4999, 0, 0, '0', ''], [0.5, 0, 0, '1', ''],
  ] as const)('formats the magnitude of %s at %s places with scale %s', (value, places, scale, integer, fraction) => {
    expect(decimalMagnitude(value, places, scale)).toEqual({ integer, fraction });
  });
  it('scales huge and tiny finite numbers without floating overflow or giant zero allocations', () => {
    expect(decimalMagnitude(1e308, 2, 2)).toEqual({ integer: '1' + '0'.repeat(310), fraction: '00' });
    expect(decimalMagnitude(Number.MIN_VALUE, 2)).toEqual({ integer: '0', fraction: '00' });
    expect(decimalMagnitude(Number.MIN_VALUE, 0, 324)).toEqual({ integer: '5', fraction: '' });
    expect(decimalMagnitude(123, 2, -8192)).toEqual({ integer: '0', fraction: '00' });
  });
  it('bounds output precision and rejects invalid numeric inputs', () => {
    for (const value of [NaN, Infinity, -Infinity]) expect(() => decimalMagnitude(value, 2)).toThrow();
    for (const places of [-1, 1.5, 2049]) expect(() => decimalMagnitude(1, places)).toThrow();
    expect(() => decimalMagnitude(1e308, 2, 8192)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
});
