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
