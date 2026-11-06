import { expect, it } from "vitest";
import { createSharedFormulaTable } from "../src/workbook/shared-formula.js";

it("expands shared followers from a per-sheet origin and retains mixed coordinates", () => {
  const groups = createSharedFormulaTable(), origin = { row: 1, column: 1 }, range = { from: { row: 1, column: 1 }, to: { row: 4, column: 2 } };
  groups.define(0, origin, range, "=A2*$D$1+Other!B$1");
  origin.row = 100; range.to.row = 100;
  expect(groups.resolve(0, { row: 1, column: 1 })).toBe("=A2*$D$1+Other!B$1");
  expect(groups.resolve(0, { row: 4, column: 2 })).toBe("=((B5*$D$1)+'Other'!C$1)");
  expect(groups.size).toBe(1);
  expect(() => groups.resolve(0, { row: 100, column: 1 })).toThrow();
  expect(() => createSharedFormulaTable().resolve(0, { row: 1, column: 1 })).toThrow();
});

it("rejects duplicate/orphan/out-of-range groups and leaves the table unchanged on budget failure", () => {
  const groups = createSharedFormulaTable(1), origin = { row: 0, column: 0 }, range = { from: origin, to: origin };
  groups.define(0, origin, range, "=1");
  expect(() => groups.define(0, origin, range, "=2")).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  expect(() => groups.define(1, origin, range, "=2")).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  expect(groups.resolve(0, origin)).toBe("=1"); expect(groups.size).toBe(1);
  expect(() => createSharedFormulaTable().define(0, { row: 1, column: 0 }, range, "=1")).toThrow();
  expect(() => groups.resolve(-1, origin)).toThrow();
  expect(() => createSharedFormulaTable(10, 1).define(0, origin, range, "=1")).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
});
