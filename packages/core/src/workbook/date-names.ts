import { freeze, unsupported } from "../document/validation.js";
import type { WorkbookDisplayLocale } from "./display-locale.js";

export interface WorkbookDateNames {
  readonly months: { readonly short: readonly string[]; readonly long: readonly string[]; readonly initial: readonly string[] };
  readonly contextualMonths: { readonly short: readonly string[]; readonly long: readonly string[] };
  readonly weekdays: { readonly short: readonly string[]; readonly long: readonly string[] };
  readonly periods: readonly [string, string];
}

/** Small, immutable locale vocabulary for a formatter instance. The synthetic
 * 1900 leap day is never passed through Date/Intl for calendar interpretation. */
export function workbookDateNames(display: WorkbookDisplayLocale): WorkbookDateNames {
  const requested = new Intl.Locale(display.locale);
  if (display.calendar > 2 || requested.calendar && requested.calendar !== "gregory"
    || display.numeralSystem || requested.numberingSystem && requested.numberingSystem !== "latn") unsupported("This workbook calendar or numeral system requires a separate display adapter.");
  const locale = display.calendar === 2 ? "en-US" : display.locale;
  const formatter = (options: Intl.DateTimeFormatOptions): Intl.DateTimeFormat => new Intl.DateTimeFormat(locale, {
    localeMatcher: "lookup", calendar: "gregory", numberingSystem: "latn", timeZone: "UTC", ...options,
  });
  const part = (format: Intl.DateTimeFormat, value: number, type: Intl.DateTimeFormatPartTypes): string => {
    const text = format.formatToParts(value).find(part => part.type === type)?.value;
    if (!text) unsupported("The requested locale has no supported calendar display names.");
    return text;
  };
  const months = (style: "short" | "long", contextual: boolean): readonly string[] => {
    const format = formatter({ month: style, ...(contextual ? { day: "numeric" } : {}) });
    return Array.from({ length: 12 }, (_, month) => part(format, Date.UTC(2000, month, 1), "month"));
  };
  const weekdays = (style: "short" | "long"): readonly string[] => {
    const format = formatter({ weekday: style });
    return Array.from({ length: 7 }, (_, day) => part(format, Date.UTC(2000, 0, 2 + day), "weekday"));
  };
  const long = months("long", false), segmenter = new Intl.Segmenter(locale, { granularity: "grapheme" });
  const period = formatter({ hour: "numeric", hourCycle: "h12" });
  return freeze({ months: { short: months("short", false), long, initial: long.map(name => segmenter.segment(name)[Symbol.iterator]().next().value!.segment) },
    contextualMonths: { short: months("short", true), long: months("long", true) },
    weekdays: { short: weekdays("short"), long: weekdays("long") },
    periods: [part(period, Date.UTC(2000, 0, 1, 1), "dayPeriod"), part(period, Date.UTC(2000, 0, 1, 13), "dayPeriod")] });
}
