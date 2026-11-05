import { describe, expect, it } from "vitest";
import { chartAxisTicks } from "../src/chart-axis.js";

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
    for (const input of [[0, 1, .1, 9], [0, 1, 0, 100], [0, Infinity, 1, 100], [-1e308, 1e308, 1e307, 100], [1, 1, 1, 100], [0, 1, NaN, 100]]) {
      expect(() => chartAxisTicks(input[0]!, input[1]!, input[2]!, input[3]!)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
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
