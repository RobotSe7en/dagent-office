import { checkLimit, invalid, unsupported } from "../document/validation.js";
import type { NumberSeparators } from "./format-numeric.js";
import { nativeWorkbookLocales } from "./native-locales.js";

export interface WorkbookLocaleOptions {
  readonly formatLocale?: number;
  readonly hostLocale?: string;
  readonly documentLocale?: string;
}
export interface WorkbookDisplayLocale {
  readonly locale: string;
  readonly source: "format" | "host" | "document" | "default";
  readonly systemFormat?: "longDate" | "time";
  /** The two upper native bytes are separate from the language ID. Retaining
   * them here does not claim that calendar/numeral conversion is implemented. */
  readonly calendar: number;
  readonly numeralSystem: number;
  readonly parseCalendar: boolean;
  readonly parseNumerals: boolean;
}

/** Resolve explicitly, including native system formats. Neither a missing tag
 * nor an unsupported tag is passed to Intl as an implicit machine default. */
export function resolveWorkbookDisplayLocale(options: WorkbookLocaleOptions = {}): WorkbookDisplayLocale {
  const native = options.formatLocale ?? 0;
  if (!Number.isSafeInteger(native) || native < 0 || native > 0xffffffff) invalid("A native format locale must be an unsigned 32-bit value.");
  const language = native & 0xffff, calendarByte = native >>> 16 & 0xff, numeralByte = native >>> 24;
  const systemFormat = language === 0xf800 ? "longDate" : language === 0xf400 ? "time" : undefined;
  let tag: string, source: WorkbookDisplayLocale["source"];
  if (language && !systemFormat) {
    tag = nativeWorkbookLocales[language]!;
    if (!tag) unsupported("The native number-format language ID has no supported assigned language tag.");
    source = "format";
  } else if (options.hostLocale !== undefined) { tag = options.hostLocale; source = "host"; }
  else if (options.documentLocale !== undefined) { tag = options.documentLocale; source = "document"; }
  else { tag = "zh-CN"; source = "default"; }
  if (typeof tag !== "string" || !tag) invalid("A workbook display locale requires a nonempty language tag.");
  checkLimit(tag.length, 255);
  let locale: string;
  try { locale = Intl.getCanonicalLocales(tag)[0]!; }
  catch { invalid("A workbook display locale must be a valid language tag."); }
  if (!Intl.NumberFormat.supportedLocalesOf([locale], { localeMatcher: "lookup" }).length
    || !Intl.DateTimeFormat.supportedLocalesOf([locale], { localeMatcher: "lookup" }).length) unsupported("This runtime does not provide the requested workbook display locale.");
  return Object.freeze({ locale, source, ...(systemFormat ? { systemFormat } : {}), calendar: calendarByte & 0x7f,
    numeralSystem: numeralByte & 0x7f, parseCalendar: Boolean(calendarByte & 0x80), parseNumerals: Boolean(numeralByte & 0x80) });
}

/** Ordinary native 0/#/? fields use Latin digits. Native numeral-system flags
 * and explicit Unicode extensions must not be silently erased by that branch. */
export function workbookLocaleSeparators(display: WorkbookDisplayLocale): NumberSeparators {
  const numbering = new Intl.Locale(display.locale).numberingSystem;
  if (display.numeralSystem || numbering && numbering !== "latn") unsupported("This workbook numeral system requires a separate display adapter.");
  const parts = new Intl.NumberFormat(display.locale, { localeMatcher: "lookup", numberingSystem: "latn", useGrouping: true,
    minimumFractionDigits: 1, maximumFractionDigits: 1 }).formatToParts(1234567890.5);
  const decimal = parts.find(part => part.type === "decimal")?.value, group = parts.find(part => part.type === "group")?.value;
  const integers = parts.filter(part => part.type === "integer").map(part => part.value.length);
  if (!decimal || !group || integers.length < 2) unsupported("The requested locale has no supported positional numeric grouping.");
  return Object.freeze({ decimal, group, primaryGroupSize: integers.at(-1)!, secondaryGroupSize: integers.at(-2)! });
}
