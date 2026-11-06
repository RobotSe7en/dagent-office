import { describe, expect, it } from "vitest";
import { chartAxisTicks, chartValueFraction, chartValueScale } from "../src/chart-axis.js";

describe("finite bounded chart ticks", () => {
  it("keeps decimal endpoints without accumulation drift", () => {
    expect(chartAxisTicks(0, 1, .1, 10)).toHaveLength(11);
    expect(chartAxisTicks(0, 1, .1, 10).at(-1)).toBe(1);
    expect(chartAxisTicks(-1, 1, .25, 8)).toEqual([-1, -.75, -.5, -.25, 0, .25, .5, .75, 1]);
    expect(chartAxisTicks(0, .3, .1, 3, .001)).toHaveLength(4);
  });
  it("rejects positive units that cannot advance at the source magnitude", () => {
    for (const [low, high, unit] of [[1e16, 1e16 + 4, .1], [-1e16, -1e16 + 4, .1], [1e308, 1.0000000000000002e308, 1e290]]) {
      expect(() => chartAxisTicks(low!, high!, unit!, 4096)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    }
  });
  it("admits finite subnormal steps and exact budgets but rejects overflow and excess ticks", () => {
    expect(chartAxisTicks(0, Number.MIN_VALUE * 3, Number.MIN_VALUE, 3)).toEqual([0, Number.MIN_VALUE, Number.MIN_VALUE * 2, Number.MIN_VALUE * 3]);
    for (const input of [[0, 1, .1, 9], [0, 1, 0, 100], [0, Infinity, 1, 100], [-1e308, 1e308, 1e307, 19], [1, 1, 1, 100], [0, 1, NaN, 100]]) {
      expect(() => chartAxisTicks(input[0]!, input[1]!, input[2]!, input[3]!)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    }
  });
  it('normalizes and steps across finite endpoints whose difference or indexed offset overflows', () => {
    const maximum = Number.MAX_VALUE;
    expect(chartAxisTicks(-1e308, 1e308, 1e307, 20)).toHaveLength(21);
    expect(chartAxisTicks(-maximum, maximum, maximum, 2)).toEqual([-maximum, 0, maximum]);
    expect([-maximum, 0, maximum].map(value => chartValueFraction(value, -maximum, maximum))).toEqual([0, .5, 1]);
    expect([-maximum, 0, maximum].map(value => chartValueFraction(value, -maximum, maximum, true))).toEqual([1, .5, 0]);
    expect(chartValueFraction(1e308, -1e308, -9e307)).toBeCloseTo(20);
    const chart = { kind: 'line' as const, legend: 'none' as const, showLabels: false };
    expect(chartValueScale({ ...chart, minimum: 1e308, majorUnit: 1e308 }, [{ values: [1e308, 1.7e308] }])).toEqual({ minimum: 1e308, maximum, unit: 1e308 });
    for (const values of [[-1e308, 1e308], [maximum], [Number.MIN_VALUE], [-Number.MIN_VALUE, Number.MIN_VALUE]]) {
      const scale = chartValueScale(chart, [{ values }]);
      expect([scale.minimum, scale.maximum, scale.unit].every(Number.isFinite)).toBe(true);
      expect(scale.maximum).toBeGreaterThan(scale.minimum); expect(scale.unit).toBeGreaterThan(0);
      expect(chartAxisTicks(scale.minimum, scale.maximum, scale.unit, 1024).length).toBeGreaterThan(1);
      for (const value of values) { expect(chartValueFraction(value, scale.minimum, scale.maximum)).toBeGreaterThanOrEqual(0); expect(chartValueFraction(value, scale.minimum, scale.maximum)).toBeLessThanOrEqual(1); }
    }
  });
});

describe('shared category intervals', () => {
  it('projects source points and grid boundaries for between and mid-category crossing', async () => {
    const { chartCategoryAxis } = await import('../src/chart-axis.js');
    for (const reverse of [false, true]) {
      const between = chartCategoryAxis(3, 10, 120, true, reverse), mid = chartCategoryAxis(3, 10, 120, false, reverse);
      expect([0, 1, 2].map(between.position)).toEqual(reverse ? [110, 70, 30] : [30, 70, 110]);
      expect([0, 1, 2].map(mid.position)).toEqual(reverse ? [130, 70, 10] : [10, 70, 130]);
      expect(Array.from({ length: between.tickCount }, (_, i) => between.tick(i))).toEqual([10, 50, 90, 130]);
      expect(Array.from({ length: mid.tickCount }, (_, i) => mid.tick(i))).toEqual([10, 70, 130]);
    }
  });
  it('keeps singleton and large source axes finite without allocating per-axis arrays', async () => {
    const { chartCategoryAxis } = await import('../src/chart-axis.js');
    const single = chartCategoryAxis(1, 10, 120, false, false); expect(single.position(0)).toBe(10); expect(single.tickCount).toBe(2); expect(single.tick(1)).toBe(130);
    expect(chartCategoryAxis(1, 10, 120, true, true).position(0)).toBe(70);
    const large = chartCategoryAxis(0xffff_ffff, 0, 1000, true, false); expect(large.tickCount).toBe(0x1_0000_0000); expect(large.tick(0xffff_ffff)).toBe(1000);
    for (const count of [0, -1, 1.5, Infinity, 0x1_0000_0000]) expect(() => chartCategoryAxis(count, 0, 120, true, false)).toThrowError(expect.objectContaining({ code: 'ResourceLimitExceeded' }));
  });
});

describe('Office axis placement', () => {
  const style = (values: Record<string, string> = {}) => ({ getPropertyValue: (name: string) => values[name] ?? '' }) as CSSStyleDeclaration;
  it('crosses on the perpendicular scale, clamps to its bounds and ignores axPos', async () => {
    const { chartAxisPlacement } = await import('../src/chart-axis.js');
    const chart = { kind: 'line' as const, legend: 'none' as const, showLabels: false };
    const category = style({ '--office-axis-crosses-at': '-2', '--office-axis-position': 't' }), value = style({ '--office-axis-crosses-at': '2', '--office-axis-position': 'r' });
    expect(chartAxisPlacement(chart, category, value, -4, 4, 3)).toMatchObject({ x: 1 / 3, y: .75, categoryTop: false, valueRight: false });
    expect(chartAxisPlacement({ ...chart, reverseCategories: true, reverseValues: true }, category, value, -4, 4, 3)).toMatchObject({ x: 1 - 1 / 3, y: .25, categoryTop: true, valueRight: true });
    expect(chartAxisPlacement(chart, style({ '--office-axis-crosses-at': '1e308' }), style({ '--office-axis-crosses-at': '-1e308' }), -4, 4, 3)).toMatchObject({ x: 0, y: 0, categoryTop: true, valueRight: false });
    expect(chartAxisPlacement(chart, style(), style(), 1, 4, 3).y).toBe(1);
    expect(chartAxisPlacement(chart, style(), style(), -4, -1, 3).y).toBe(0);
  });
  it('keeps labels at high/low ends independently from the axis and reverses their sides', async () => {
    const { chartAxisPlacement } = await import('../src/chart-axis.js');
    const chart = { kind: 'column' as const, legend: 'none' as const, showLabels: false };
    for (const reverse of [false, true]) {
      const at = chartAxisPlacement({ ...chart, reverseCategories: reverse, reverseValues: reverse }, style({ '--office-axis-label-position': 'high', '--office-axis-crosses-at': '2' }), style({ '--office-axis-label-position': 'low', '--office-axis-crosses-at': '2', '--office-axis-cross-between': 'midCat' }), -4, 4, 3);
      expect(at).toMatchObject({ x: .5, y: reverse ? .75 : .25, categoryLabel: reverse ? 1 : 0, valueLabel: reverse ? 1 : 0, categoryTop: !reverse, valueRight: reverse });
    }
  });
  it('rejects nonfinite, nonintegral category crossings and malformed persisted CSS', async () => {
    const { chartAxisPlacement } = await import('../src/chart-axis.js');
    const chart = { kind: 'line' as const, legend: 'none' as const, showLabels: false };
    for (const invalid of ['NaN', 'Infinity', '0x10', '2.5', '1e309']) expect(() => chartAxisPlacement(chart, style(), style({ '--office-axis-crosses-at': invalid }), -4, 4, 3)).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
    expect(() => chartAxisPlacement(chart, style({ '--office-axis-crosses': 'minimum' }), style(), -4, 4, 3)).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  });
});
