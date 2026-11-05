import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import type { WorkbookDateSystem } from "../src/workbook/calendar.js";
import { workbookDateNames } from "../src/workbook/date-names.js";
import { resolveWorkbookDisplayLocale, workbookLocaleSeparators } from "../src/workbook/display-locale.js";
import { formatDateSection } from "../src/workbook/format-date.js";
import { compileNumberFormat } from "../src/workbook/number-format.js";

const en = resolveWorkbookDisplayLocale({ hostLocale: "en-US" }), names = workbookDateNames(en);
const display = (value: number, code: string, system: WorkbookDateSystem = "1900") => formatDateSection(value, compileNumberFormat(code)[0]!, system, names, ".");
const text = (value: number, code: string, system: WorkbookDateSystem = "1900") => {
  const result = display(value, code, system);
  if (result.kind !== "parts") throw new Error("Expected a date display");
  return result.parts.map(part => part.kind === "text" ? part.text : `<${part.kind}>`).join("");
};

describe("workbook calendar and time display", () => {
  it("formats synthetic 1900 dates, weekdays and the distinct 1904 base without Date normalization", () => {
    for (const [value, expected] of [[0, "1900-01-00 Sat"], [1, "1900-01-01 Sun"], [59, "1900-02-28 Tue"],
      [60, "1900-02-29 Wed"], [61, "1900-03-01 Thu"], [45351, "2024-02-29 Thu"], [2958465, "9999-12-31 Fri"]] as const) expect(text(value, "yyyy-mm-dd ddd")).toBe(expected);
    expect(text(0, "yyyy-mm-dd ddd", "1904")).toBe("1904-01-01 Fri");
    expect(text(59, "yyyy-mm-dd ddd", "1904")).toBe("1904-02-29 Mon");
    expect(text(2957003, "yyyy-mm-dd ddd", "1904")).toBe("9999-12-31 Fri");
    expect(text(45351.5, 'dddd, mmmm d, yyyy "at" h:mm AM/PM')).toBe("Thursday, February 29, 2024 at 12:00 PM");
    expect(text(45351.5, "mm h:mm mm:ss yy mmmmm")).toBe("02 12:00 00:00 24 F");
    expect(text(0, "h:mm am/pm")).toBe("12:00 am");
    expect(text(.75, "h:mm A/P")).toBe("6:00 P");
  });
  it("rounds displayed time before omitting units and carries through seconds, days and the leap compatibility boundary", () => {
    for (const [seconds, expected] of [[.5, "00:00:01"], [1.5, "00:00:02"], [29.5, "00:00:30"], [59.5, "00:01:00"], [3599.5, "01:00:00"]] as const) expect(text(seconds / 86400, "hh:mm:ss")).toBe(expected);
    expect(text(.499999 / 86400, "hh:mm:ss")).toBe("00:00:00");
    expect(text(.500001 / 86400, "hh:mm:ss")).toBe("00:00:01");
    expect(text(.0005 / 86400, "hh:mm:ss.000")).toBe("00:00:00.001");
    expect(text(59.9 / 86400, "hh:mm:ss")).toBe("00:01:00");
    expect(text(59.9 / 86400, "hh:mm")).toBe("00:01");
    expect(text(29.9 / 86400, "hh:mm")).toBe("00:00");
    expect(text(59.9 / 86400, "hh:mm:ss.00")).toBe("00:00:59.90");
    expect(text(59.9996 / 86400, "hh:mm:ss.000")).toBe("00:01:00.000");
    expect(text(60 - .0004 / 86400, "yyyy-mm-dd hh:mm:ss.000")).toBe("1900-02-29 00:00:00.000");
    expect(text(61 - .1 / 86400, "yyyy-mm-dd hh:mm:ss")).toBe("1900-03-01 00:00:00");
    expect(text(61 - .1 / 86400, "yyyy-mm-dd hh:mm:ss.0")).toBe("1900-02-29 23:59:59.9");
    expect(text(1 - .0004 / 86400, "yyyy-mm-dd")).toBe("1900-01-01");
    expect(text(1 - .001 / 86400, "yyyy-mm-dd")).toBe("1900-01-00");
    expect(display(2958466 - .1 / 86400, "yyyy-mm-dd hh:mm:ss")).toEqual({ kind: "invalidDate" });
    expect(text(Number.MIN_VALUE, "hh:mm:ss.000")).toBe("00:00:00.000");
  });
  it("formats elapsed time without wrapping at a day or modifying the original scalar", () => {
    const value = 1.25;
    expect(text(value, "[h]:mm:ss")).toBe("30:00:00");
    expect(text(value, "[mm]:ss")).toBe("1800:00");
    expect(text(value, "[ss].00")).toBe("108000.00");
    expect(text(.25, "[hh]:mm")).toBe("06:00");
    expect(value).toBe(1.25);
    const section = compileNumberFormat('yyyy-mm-dd_)*禰');
    const result = formatDateSection(60, section[0]!, "1900", names, ".");
    expect(result).toEqual({ kind: "parts", parts: [{ kind: "text", text: "1900-02-29" }, { kind: "space", character: ")" }, { kind: "fill", character: "禰" }] });
    expect(Object.isFrozen(result)).toBe(true);
    expect(compileNumberFormat('yyyy-mm-dd_)*禰')).toEqual(section);
  });
  it("matches the compatible date and elapsed cases in the pinned independent producer evidence", () => {
    const reference = JSON.parse(readFileSync("docs/evidence/xlsx-number-format-semantics.json", "utf8")) as { cases: { dateSystem: WorkbookDateSystem; samples: { value: number; code: string; text: string }[] }[] };
    let checked = 0;
    for (const series of reference.cases) for (const sample of series.samples.slice(0, 10)) {
      // Calc's early 1900 calendar and whole-second truncation differ from the
      // chosen Excel semantics; preserve those recorded observations separately.
      if (sample.code === "yyyy-mm-dd" && series.dateSystem === "1900" && sample.value < 61 || sample.code === "hh:mm" || sample.code === "hh:mm:ss") continue;
      expect(text(sample.value, sample.code, series.dateSystem)).toBe(sample.text); checked++;
    }
    expect(checked).toBe(12);
  });
  it("uses locale month forms and explicit Gregorian calendar semantics", () => {
    const fr = resolveWorkbookDisplayLocale({ hostLocale: "fr-FR" }), ru = resolveWorkbookDisplayLocale({ hostLocale: "ru-RU" });
    const translated = (locale: typeof fr, code: string) => formatDateSection(45351.5, compileNumberFormat(code)[0]!, "1900", workbookDateNames(locale), workbookLocaleSeparators(locale).decimal);
    expect(translated(fr, "dddd d mmmm yyyy hh:mm:ss.00")).toEqual({ kind: "parts", parts: [{ kind: "text", text: "jeudi 29 février 2024 12:00:00,00" }] });
    expect(translated(ru, "d mmmm yyyy")).toEqual({ kind: "parts", parts: [{ kind: "text", text: "29 февраля 2024" }] });
    expect(workbookDateNames(resolveWorkbookDisplayLocale({ hostLocale: "zh-CN" })).months.long[1]).toBe("二月");
    expect(workbookDateNames(resolveWorkbookDisplayLocale({ formatLocale: 0x20407 })).months.long[1]).toBe("February");
    for (const options of [{ formatLocale: 0x30411 }, { hostLocale: "th-TH-u-ca-buddhist" }, { hostLocale: "ar-EG-u-nu-arab" }]) expect(() => workbookDateNames(resolveWorkbookDisplayLocale(options))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
  it("distinguishes invalid display values from unsupported source grammar", () => {
    for (const value of [-1, -Number.MIN_VALUE, 2958466, Number.MAX_VALUE]) expect(display(value, "yyyy-mm-dd")).toEqual({ kind: "invalidDate" });
    for (const value of [NaN, Infinity, -Infinity]) expect(() => display(value, "yyyy-mm-dd")).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    for (const code of ["y", "yyy", "dddddd", "hhh", "mmmmmm", "ss.0000", "ss.#", "ss.0.00", "hh.00", "[h]:[m]", "[h] h", "[h] yyyy", "0 yyyy", "ss.00%", "General", "[$-F800]yyyy-mm-dd"]) {
      expect(() => display(1, code), code).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    }
  });
});
