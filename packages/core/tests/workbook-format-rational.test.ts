import { describe, expect, it } from "vitest";
import { nearestDisplayFraction, rationalMagnitude } from "../src/workbook/format-rational.js";

describe('bounded rational magnitude for numeric display', () => {
  it('keeps decimal scaling exact without multiplying a finite scalar into infinity', () => {
    expect(rationalMagnitude(-1.3)).toEqual({ numerator: 13n, denominator: 10n });
    expect(rationalMagnitude(1e308, 2)).toEqual({ numerator: 10n ** 310n, denominator: 1n });
    expect(rationalMagnitude(Number.MIN_VALUE)).toEqual({ numerator: 5n, denominator: 10n ** 324n });
  });
  it('matches independent standard fraction examples and native one/two-digit denominator probes', () => {
    expect(nearestDisplayFraction(rationalMagnitude(1.3), 9n)).toEqual({ numerator: 9n, denominator: 7n });
    expect(nearestDisplayFraction(rationalMagnitude(1.3), 99n)).toEqual({ numerator: 13n, denominator: 10n });
    expect(nearestDisplayFraction(rationalMagnitude(1.99), 99n)).toEqual({ numerator: 197n, denominator: 99n });
    expect(nearestDisplayFraction(rationalMagnitude(Math.PI), 999n)).toEqual({ numerator: 355n, denominator: 113n });
    expect(nearestDisplayFraction(rationalMagnitude(0.5), 1n)).toEqual({ numerator: 0n, denominator: 1n });
    expect(nearestDisplayFraction(rationalMagnitude(0), 999n)).toEqual({ numerator: 0n, denominator: 1n });
  });
  it('agrees with exhaustive independent denominator enumeration on a small rational grid', () => {
    const abs = (value: bigint) => value < 0n ? -value : value;
    for (let n = 0n; n <= 254n; n++) {
      const source = { numerator: n, denominator: 127n }, selected = nearestDisplayFraction(source, 19n);
      const distance = abs(n * selected.denominator - selected.numerator * 127n);
      for (let d = 1n; d <= 19n; d++) {
        const lower = n * d / 127n;
        for (const p of [lower, lower + 1n]) expect(distance * d <= abs(n * d - p * 127n) * selected.denominator).toBe(true);
      }
    }
  });
  it('handles enormous denominator bounds without enumeration and rejects invalid budgets', () => {
    expect(nearestDisplayFraction(rationalMagnitude(0.125), 10n ** 1000n)).toEqual({ numerator: 1n, denominator: 8n });
    expect(() => nearestDisplayFraction({ numerator: -1n, denominator: 1n }, 9n)).toThrow();
    expect(() => nearestDisplayFraction({ numerator: 1n, denominator: 0n }, 9n)).toThrow();
    expect(() => nearestDisplayFraction({ numerator: 1n, denominator: 1n }, 0n)).toThrow();
    expect(() => rationalMagnitude(1e308, 8192)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
});
