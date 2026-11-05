import { describe, expect, it } from "vitest";
import { resolveWorkbookDisplayLocale, workbookLocaleSeparators } from "../src/workbook/display-locale.js";
import { formatWorkbookValue } from "../src/workbook/format-value.js";
import { compileNumberFormat } from "../src/workbook/number-format.js";
import { nativeWorkbookLocales } from "../src/workbook/native-locales.js";

describe("explicit workbook display locales", () => {
  it("selects format, host, document and zh-CN in that order without changing the supplied tag", () => {
    const options = Object.freeze({ formatLocale: 0x409, hostLocale: "de-DE", documentLocale: "fr-FR" });
    expect(resolveWorkbookDisplayLocale(options)).toMatchObject({ locale: "en-US", source: "format" });
    expect(resolveWorkbookDisplayLocale({ hostLocale: "de-de", documentLocale: "fr-FR" })).toMatchObject({ locale: "de-DE", source: "host" });
    expect(resolveWorkbookDisplayLocale({ documentLocale: "fr-fr" })).toMatchObject({ locale: "fr-FR", source: "document" });
    expect(resolveWorkbookDisplayLocale()).toMatchObject({ locale: "zh-CN", source: "default" });
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0x804, hostLocale: "invalid_tag" })).toMatchObject({ locale: "zh-CN", source: "format" });
    expect(Object.isFrozen(resolveWorkbookDisplayLocale(options))).toBe(true);
    expect(options).toEqual({ formatLocale: 0x409, hostLocale: "de-DE", documentLocale: "fr-FR" });
  });
  it("keeps native system formats and high calendar/numeral bytes separate from language selection", () => {
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0xf800, hostLocale: "de-DE" })).toMatchObject({ locale: "de-DE", source: "host", systemFormat: "longDate" });
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0xf400 })).toMatchObject({ locale: "zh-CN", source: "default", systemFormat: "time" });
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0x82810409 })).toEqual({ locale: "en-US", source: "format", calendar: 1, numeralSystem: 2, parseCalendar: true, parseNumerals: true });
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0x40a })).toMatchObject({ locale: "es-ES" });
    expect(resolveWorkbookDisplayLocale({ formatLocale: 0 })).toMatchObject({ locale: "zh-CN" });
  });
  it("uses regional separators and grouping with real scalar formatting", () => {
    const values = [
      ["en-US", "1,234,567.50"], ["de-DE", "1.234.567,50"], ["fr-FR", "1\u202f234\u202f567,50"],
      ["hi-IN", "12,34,567.50"], ["zh-CN", "1,234,567.50"], ["de-CH", "1’234’567.50"],
    ] as const;
    for (const [hostLocale, expected] of values) {
      const separators = workbookLocaleSeparators(resolveWorkbookDisplayLocale({ hostLocale }));
      expect(formatWorkbookValue(1234567.5, compileNumberFormat("#,##0.00"), separators)).toEqual({ kind: "parts", parts: [{ kind: "text", text: expected }] });
      expect(Object.isFrozen(separators)).toBe(true);
    }
  });
  it("never falls back to the machine locale for invalid, unknown or unsupported explicit choices", () => {
    for (const hostLocale of ["", "en_US", "de--DE"]) expect(() => resolveWorkbookDisplayLocale({ hostLocale, documentLocale: "zh-CN" })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    for (const hostLocale of ["zz-ZZ", "qps-ploc"]) expect(() => resolveWorkbookDisplayLocale({ hostLocale })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    for (const formatLocale of [0x7f, 0x1000, 0xffff]) expect(() => resolveWorkbookDisplayLocale({ formatLocale, hostLocale: "zh-CN" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    for (const formatLocale of [-1, NaN, Infinity, .5, 0x100000000]) expect(() => resolveWorkbookDisplayLocale({ formatLocale })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => resolveWorkbookDisplayLocale({ hostLocale: "x".repeat(256) })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    for (const options of [{ formatLocale: 0x01000409 }, { hostLocale: "en-US-u-nu-arab" }]) expect(() => workbookLocaleSeparators(resolveWorkbookDisplayLocale(options))).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(Object.keys(nativeWorkbookLocales)).toHaveLength(412);
    for (const tag of Object.values(nativeWorkbookLocales)) expect(Intl.getCanonicalLocales(tag)).toHaveLength(1);
  });
});
