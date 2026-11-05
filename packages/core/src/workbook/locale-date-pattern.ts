import { unsupported } from "../document/validation.js";
import type { WorkbookDisplayLocale } from "./display-locale.js";

/** Locale data supplies a derived pattern, never a replacement persisted
 * formatCode. Escape every literal so localized words cannot become fields. */
export function workbookLocaleDatePattern(display: WorkbookDisplayLocale, kind: "shortDate" | "longDate" | "time"): string {
  const requested = new Intl.Locale(display.locale);
  if (display.calendar > 2 || requested.calendar && requested.calendar !== "gregory"
    || display.numeralSystem || requested.numberingSystem && requested.numberingSystem !== "latn") unsupported("This locale date pattern requires a separate calendar or numeral adapter.");
  const formatter = new Intl.DateTimeFormat(display.calendar === 2 ? "en-US" : display.locale, {
    calendar: "gregory", numberingSystem: "latn", localeMatcher: "lookup", timeZone: "UTC",
    ...(kind === "time" ? { timeStyle: "medium" } : { dateStyle: kind === "shortDate" ? "short" : "full" }),
  });
  if (kind === "time" && ["h11", "h24"].includes(formatter.resolvedOptions().hourCycle ?? "")) unsupported("This locale hour cycle cannot be represented by native h/hh fields.");
  const escape = (text: string): string => [...text].map(character => "\\" + character).join("");
  // Single-digit month/day expose the locale's leading-zero convention. A
  // modern year avoids era/year-window ambiguities in the pattern probe.
  const parts = formatter.formatToParts(Date.UTC(2001, 1, 3, 13, 4, 5));
  return parts.map((part, index) => {
    switch (part.type) {
      case "year": return "yyyy";
      case "month": return /^\d+$/.test(part.value) ? part.value.length === 2 ? "mm" : "m" : "mmmm";
      case "day": return part.value.length === 2 ? "dd" : "d";
      case "weekday": return "dddd";
      case "hour": return part.value.length === 2 ? "hh" : "h";
      case "minute": return "mm";
      case "second": return "ss";
      case "dayPeriod": return "AM/PM";
      // ICU versions differ on the space before AM/PM. Normalize only this
      // derived separator; source literals and numeric grouping stay intact.
      case "literal": return escape(kind === "time" && (parts[index - 1]?.type === "dayPeriod" || parts[index + 1]?.type === "dayPeriod")
        ? part.value.replace(/[\u00a0\u202f]/g, " ") : part.value);
      default: return unsupported("This locale date pattern contains an unsupported field.");
    }
  }).join("");
}
