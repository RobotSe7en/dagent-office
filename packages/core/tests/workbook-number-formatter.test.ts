import { describe, expect, it } from "vitest";
import { readFileSync } from "node:fs";
import { createWorkbookNumberFormatter } from "@doffi/core";
import type { WorkbookDisplayFormat, WorkbookNumberDisplay, WorkbookNumberFormatterOptions } from "@doffi/core";

function content(display: WorkbookNumberDisplay): string {
  if (display.kind !== "parts") throw new Error("Expected display parts");
  return display.parts.map(part => part.kind === "text" ? part.text : `<${part.kind}:${part.character}>`).join("");
}
describe("public workbook number formatter", () => {
  it("maps text fields by format structure, including identical affixes and repeated UTF-16 text", () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
    expect(formatter.format("ID ", { code: '"ID "@' })).toMatchObject({ kind: "parts", parts: [
      { kind: "text", text: "ID " }, { kind: "text", text: "ID ", source: { from: 0, to: 3 } },
    ] });
    const repeated = formatter.format("😀禰󠄀", { code: '@"/"@' });
    expect(repeated).toMatchObject({ kind: "parts", parts: [
      { kind: "text", text: "😀禰󠄀", source: { from: 0, to: 5 } }, { kind: "text", text: "/" },
      { kind: "text", text: "😀禰󠄀", source: { from: 0, to: 5 } },
    ] });
    if (repeated.kind !== "parts") throw new Error("Expected text parts");
    expect(repeated.parts[1]).not.toHaveProperty("source"); expect(Object.isFrozen(repeated.parts[0])).toBe(true);
    expect(formatter.format("001", { code: "0.00" })).toMatchObject({ parts: [{ kind: "text", text: "001", source: { from: 0, to: 3 } }] });
    expect(formatter.format("", { code: '"prefix"@' })).toMatchObject({ parts: [{ kind: "text", text: "prefix" }, { kind: "text", text: "", source: { from: 0, to: 0 } }] });
    expect(formatter.format("hidden", { code: ";;;" })).toMatchObject({ parts: [] });
    expect(() => formatter.format("x".repeat(20000), { code: "@@" })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    formatter.dispose();
  });
  it("matches native plain text-section observations without treating Calc's rich-text policy as acceptance", () => {
    const evidence = JSON.parse(readFileSync("docs/evidence/xlsx-text-formats.json", "utf8")) as { samples: { value: string; code: string; text: string; rich?: boolean }[] };
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
    const cases = evidence.samples.filter(sample => !sample.rich && !/[*_]/.test(sample.code));
    expect(cases).toHaveLength(6);
    for (const sample of cases) expect(content(formatter.format(sample.value, { code: sample.code })), sample.code).toBe(sample.text);
    formatter.dispose();
  });
  it("uses current values and the selected section's locale, preserving source types and native definitions", () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "de-DE", documentLocale: "fr-FR" });
    const definition = Object.freeze({ formatId: "currency", code: '#,##0.00;[Red][$¥-804](#,##0.00)' });
    expect(formatter.format(1234.5, definition)).toMatchObject({ locale: "de-DE", localeSource: "host" });
    expect(content(formatter.format(1234.5, definition))).toBe("1.234,50");
    expect(formatter.format(-1234.5, definition)).toEqual({ kind: "parts", locale: "zh-CN", localeSource: "format", color: { kind: "named", name: "red" }, parts: [{ kind: "text", text: "¥(1,234.50)" }] });
    expect(content(formatter.format("00123", definition))).toBe("00123");
    expect(content(formatter.format("00123", { code: '0;0;0;"ID "@' }))).toBe("ID 00123");
    expect(content(formatter.format(.125, { builtinId: 10 }))).toBe("12,50%");
    expect(content(formatter.format(1.25, { builtinId: 12 }))).toBe("1 1/4");
    expect(definition).toEqual({ formatId: "currency", code: '#,##0.00;[Red][$¥-804](#,##0.00)' });
    expect(Object.isFrozen(formatter.format(1))).toBe(true);
  });
  it("resolves built-in short dates at display time and keeps Excel's compatibility day", () => {
    const expected = [["zh-CN", "1900/2/29"], ["en-US", "2/29/1900"], ["de-DE", "29.02.1900"], ["fr-FR", "29/02/1900"]] as const;
    const definition = Object.freeze({ formatId: "native-date", builtinId: 14 });
    for (const [hostLocale, text] of expected) {
      const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale });
      expect(content(formatter.format(60, definition))).toBe(text);
      expect(content(formatter.format(60.5, { builtinId: 22 }))).toBe(text + " 12:00");
      formatter.dispose();
    }
    const zh = createWorkbookNumberFormatter({ dateSystem: "1900" });
    expect(content(zh.format(60.5, { builtinId: 31 }))).toBe("1900年2月29日");
    expect(content(zh.format(60.5, { builtinId: 56 }))).toBe("下午 12时00分00秒");
    const old = createWorkbookNumberFormatter({ dateSystem: "1904", hostLocale: "en-US" });
    expect(content(old.format(0, definition))).toBe("1/1/1904");
    expect(definition).toEqual({ formatId: "native-date", builtinId: 14 });
  });
  it("replaces native system date/time patterns with the chosen locale's pattern", () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "en-US" });
    expect(content(formatter.format(60, { code: '[$-F800]"ignored"yyyy-mm-dd' }))).toBe("Wednesday, February 29, 1900");
    expect(content(formatter.format(0, { code: "[$-F800]yyyy-mm-dd" }))).toBe("Saturday, January 0, 1900");
    expect(content(formatter.format(.5, { code: "[$-F400]h:mm" }))).toBe("12:00:00 PM");
    expect(content(formatter.format(.5, { code: 'h:mm"\u202f"AM/PM' }))).toBe("12:00\u202fPM");
    expect(formatter.format(60, { code: "[Red][$-F800]yyyy-mm-dd" })).toMatchObject({ color: { kind: "named", name: "red" }, localeSource: "host" });
    const zh = createWorkbookNumberFormatter({ dateSystem: "1900" });
    expect(content(zh.format(60, { code: "[$-F800]yyyy-mm-dd" }))).toBe("1900年2月29日星期三");
    const specialClock = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "en-US-u-hc-h24" });
    expect(() => specialClock.format(0, { code: "[$-F400]h:mm" })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(content(specialClock.format(0, { code: "hh:mm" }))).toBe("00:00");
  });
  it("matches independently generated native General sections and system patterns", () => {
    const evidence = JSON.parse(readFileSync(new URL("../../../docs/evidence/xlsx-general-sections.json", import.meta.url), "utf8")) as {
      hostLocale: string; dateSystem: "1900"; samples: { value: number; code: string; text: string }[];
    };
    const formatter = createWorkbookNumberFormatter({ dateSystem: evidence.dateSystem, hostLocale: evidence.hostLocale });
    expect(evidence.samples).toHaveLength(19);
    for (const sample of evidence.samples) expect(content(formatter.format(sample.value, { code: sample.code })), sample.code).toBe(sample.text);
    formatter.dispose();
  });
  it("fits the whole General section, including its sign and literal affixes", () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "en-US" }), measurement = { availableWidth: 5, measure: (text: string) => text.length };
    expect(content(formatter.format(25.76, { code: '"ID "General' }, measurement))).toBe("ID 26");
    expect(content(formatter.format(-12.5, { code: '"ID "General' }))).toBe("-ID 12.5");
    expect(content(formatter.format(-12.5, { code: "General;(General)" }))).toBe("(12.5)");
    expect(content(formatter.format(-12.5, { code: "[<0]General;0" }))).toBe("12.5");
    expect(content(formatter.format(-12.5, { code: "General General" }))).toBe("-12.5 12.5");
    expect(content(formatter.format(-12.5, { code: "General%" }))).toBe("-1250%");
    expect(content(formatter.format(1e308, { code: "General%" }))).toBe("1E+310%");
    expect(content(formatter.format(Number.MIN_VALUE, { code: "General%" }))).toBe("5E-322%");
    expect(content(formatter.format(12.5, { code: '"$"*xGeneral_)' }))).toBe("$<fill:x>12.5<space:)>");
    expect(formatter.format(12345, undefined, { ...measurement, availableWidth: 4 })).toMatchObject({ kind: "overflow", reason: "width" });
  });
  it("distinguishes display overflow from malformed formats without returning imported caches", () => {
    const formatter = createWorkbookNumberFormatter({ dateSystem: "1900" });
    expect(formatter.format(-1, { builtinId: 14 })).toMatchObject({ kind: "overflow", reason: "invalidDate" });
    expect(formatter.format(5, { code: "[>10]0;[<0]0" })).toMatchObject({ kind: "overflow", reason: "unmatchedCondition" });
    expect(() => formatter.format(1, { builtinId: 5 })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
    expect(() => formatter.format(1, { builtinId: 164 })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => formatter.format(1, { code: "0", builtinId: 1 } as unknown as WorkbookDisplayFormat)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => formatter.format(Infinity)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => formatter.format("x".repeat(32768))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => formatter.format(1, undefined, { availableWidth: 1, measure: () => NaN })).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
    expect(() => formatter.format(1, { code: '"X"[$-F800]yyyy-mm-dd' })).toThrowError(expect.objectContaining({ code: "UnsupportedFeature" }));
  });
  it("isolates source, locale and cache lifetimes across formatters and source changes", () => {
    const options = { dateSystem: "1900", hostLocale: "de-DE" } satisfies WorkbookNumberFormatterOptions;
    const de = createWorkbookNumberFormatter(options), us = createWorkbookNumberFormatter({ dateSystem: "1900", hostLocale: "en-US" });
    options.hostLocale = "fr-FR";
    const format = { formatId: "same-id", code: "0.0" };
    expect(content(de.format(12.5, format))).toBe("12,5"); expect(content(us.format(12.5, format))).toBe("12.5");
    format.code = '"changed "0.00';
    expect(content(de.format(12.5, format))).toBe("changed 12,50");
    for (let index = 0; index < 300; index++) expect(content(de.format(index, { code: '"item ' + index + ' "0' }))).toBe(`item ${index} ${index}`);
    expect(content(de.format(12.5, { code: "0.0" }))).toBe("12,5");
    de.dispose(); de.dispose();
    expect(() => de.format(1)).toThrowError(expect.objectContaining({ code: "InvalidOperation" }));
    expect(content(us.format(12.5))).toBe("12.5");
  });
});
