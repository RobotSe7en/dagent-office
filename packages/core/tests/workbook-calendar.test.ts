import { describe, expect, it } from "vitest";
import { workbookSerialDate } from "../src/workbook/calendar.js";

describe("workbook compatibility calendars", () => {
  it("retains Excel's zero day, synthetic leap day and historical weekday sequence", () => {
    const expected = [
      [0, 1900, 1, 0, 6, "zero-day"], [1, 1900, 1, 1, 0, undefined], [31, 1900, 1, 31, 2, undefined],
      [32, 1900, 2, 1, 3, undefined], [59, 1900, 2, 28, 2, undefined], [60, 1900, 2, 29, 3, "leap-day"],
      [61, 1900, 3, 1, 4, undefined], [366, 1900, 12, 31, 1, undefined],
    ] as const;
    for (const [serial, year, month, day, weekday, compatibility] of expected) {
      const parts = workbookSerialDate(serial + .25, "1900")!;
      expect(parts).toMatchObject({ year, month, day, weekday, serialDay: serial, dayFraction: .25 });
      expect(parts.compatibility).toBe(compatibility); expect(Object.isFrozen(parts)).toBe(true);
    }
  });
  it("uses the actual Gregorian calendar after the compatibility boundary and in 1904 mode", () => {
    const examples = [
      [1462, 1904, 1, 1, 5], [1521, 1904, 2, 29, 1], [36585, 2000, 2, 29, 2],
      [45351, 2024, 2, 29, 4], [73109, 2100, 2, 28, 0], [73110, 2100, 3, 1, 1], [2958465, 9999, 12, 31, 5],
    ] as const;
    for (const [serial, year, month, day, weekday] of examples) for (const system of ["1900", "1904"] as const) {
      const native = serial - (system === "1904" ? 1462 : 0), parts = workbookSerialDate(native + .5, system)!;
      expect(parts).toEqual({ year, month, day, weekday, serialDay: native, dayFraction: .5 });
    }
    expect(workbookSerialDate(59, "1904")).toMatchObject({ year: 1904, month: 2, day: 29 });
    expect(workbookSerialDate(60, "1904")).toMatchObject({ year: 1904, month: 3, day: 1 });
  });
  it("keeps fractional serials and distinguishes invalid inputs from unrepresentable dates", () => {
    expect(workbookSerialDate(Number.MIN_VALUE, "1900")!.dayFraction).toBe(Number.MIN_VALUE);
    expect(workbookSerialDate(60.999, "1900")).toMatchObject({ day: 29, serialDay: 60 });
    expect(workbookSerialDate(-0, "1900")).toMatchObject({ day: 0 });
    for (const system of ["1900", "1904"] as const) {
      for (const value of [-Number.MIN_VALUE, -1, 2958466, Number.MAX_VALUE]) expect(workbookSerialDate(value, system)).toBeUndefined();
      for (const value of [NaN, Infinity, -Infinity]) expect(() => workbookSerialDate(value, system)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    }
    expect(workbookSerialDate(2957004, "1904")).toBeUndefined();
    expect(() => workbookSerialDate(0, "unknown" as "1900")).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
});
