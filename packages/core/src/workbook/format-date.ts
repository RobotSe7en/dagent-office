import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";
import { workbookSerialDate } from "./calendar.js";
import type { WorkbookDateSystem } from "./calendar.js";
import type { WorkbookDateNames } from "./date-names.js";
import type { NumberDisplayPart } from "./format-numeric.js";
import type { NumberFormatSection, NumberFormatToken } from "./number-format.js";

export type WorkbookDateDisplay = { readonly kind: "parts"; readonly parts: readonly NumberDisplayPart[] } | { readonly kind: "invalidDate" };
interface DatePlan { readonly minutes: ReadonlySet<number>; readonly fractional?: number; readonly precision: number; readonly time: boolean; readonly dayOfMonth: boolean }
type DateField = Extract<NumberFormatToken, { kind: "date" | "elapsed" }>;
function planDate(section: NumberFormatSection): DatePlan {
  const tokens = section.tokens, fields = tokens.flatMap((token, index) => token.kind === "date" || token.kind === "elapsed" ? [{ token, index }] : []);
  const minutes = new Set<number>(); let fractional: number | undefined, precision = 0, elapsed: DateField | undefined;
  for (let position = 0; position < fields.length; position++) {
    const { token, index } = fields[position]!;
    if (token.kind === "elapsed") { if (elapsed) unsupported("A native time section can contain only one elapsed field."); elapsed = token; continue; }
    if (token.field === "m" && token.length <= 2 && (fields[position - 1]?.token.field === "h" || fields[position + 1]?.token.field === "s")) minutes.add(index);
    const lengths = token.field === "y" ? [2, 4] : token.field === "m" ? [1, 2, 3, 4, 5] : token.field === "d" ? [1, 2, 3, 4] : [1, 2];
    if (!lengths.includes(token.length)) unsupported("This date field length is outside the native display grammar.");
  }
  if (!fields.length) unsupported("A date display section requires a calendar or time field.");
  if (elapsed && fields.some(({ token, index }) => token.kind === "date" && (token.field === elapsed!.field
    || token.field === "y" || token.field === "d" || token.field === "m" && !minutes.has(index)))) unsupported("Elapsed time cannot be mixed with calendar dates or its equivalent ordinary field.");
  for (let index = 0; index < tokens.length; index++) {
    const token = tokens[index]!;
    if (["literal", "date", "elapsed", "dayPeriod", "space", "fill", "slash", "comma"].includes(token.kind)) continue;
    if (token.kind !== "decimal" || fractional !== undefined) unsupported("This date section mixes incompatible numeric fields.");
    const previous = tokens[index - 1];
    if (!previous || previous.kind !== "date" && previous.kind !== "elapsed" || previous.field !== "s") unsupported("A fractional-second field must directly follow seconds.");
    fractional = index;
    while (tokens[index + 1]?.kind === "digit" && (tokens[index + 1] as Extract<NumberFormatToken, { kind: "digit" }>).value === "0") { precision++; index++; }
    if (precision < 1 || precision > 3) unsupported("Native fractional seconds require one to three zero placeholders.");
  }
  return { minutes, ...(fractional === undefined ? {} : { fractional }), precision,
    time: Boolean(elapsed) || minutes.size > 0 || fields.some(({ token }) => token.field === "h" || token.field === "s") || tokens.some(token => token.kind === "dayPeriod"),
    dayOfMonth: fields.some(({ token }) => token.field === "d" && token.length <= 2) };
}

/** Format one selected Gregorian/native elapsed-time section. Locale vocabulary
 * is prepared once per formatter, while the original scalar and source tokens
 * remain unchanged. Invalid display dates become width-filled hashes upstream. */
export function formatDateSection(value: number, section: NumberFormatSection, system: WorkbookDateSystem,
  names: WorkbookDateNames, decimal: string): WorkbookDateDisplay {
  if (!Number.isFinite(value) || !decimal || decimal.length > 8) invalid("A date display requires a finite scalar and a bounded decimal separator.");
  if (section.locale !== undefined && [0xf800, 0xf400].includes(section.locale & 0xffff)) unsupported("Native system date/time formats require the locale pattern adapter.");
  const plan = planDate(section);
  if (!workbookSerialDate(value, system)) return { kind: "invalidDate" };
  // Scale the native binary serial once, then round the displayed unit. A tie
  // can land one representable step below half when dividing seconds by 86400
  // and scaling back. Normalize only that immediate floating-point boundary.
  // The checked range bounds even milliseconds below 2^48, preserving integers.
  // Time fields round to seconds (or explicit subseconds), then omit unwanted
  // units; calendar-only fields normalize midnight at millisecond resolution.
  const precision = plan.time ? plan.precision : 3, unit = 10 ** precision, scaled = value * (86400 * unit);
  const floor = Math.floor(scaled), step = 2 ** (Math.floor(Math.log2(scaled)) - 52);
  const ticks = Math.abs(scaled - floor - .5) <= step ? floor + 1 : Math.round(scaled);
  const day = Math.floor(ticks / (86400 * unit)), date = workbookSerialDate(day, system);
  if (!date) return { kind: "invalidDate" };
  const seconds = Math.floor(ticks / unit), hour = Math.floor(seconds / 3600) % 24, minute = Math.floor(seconds / 60) % 60, second = seconds % 60;
  const twelveHour = section.tokens.some(token => token.kind === "dayPeriod"), parts: NumberDisplayPart[] = []; let units = 0;
  const text = (text: string): void => {
    units += text.length; checkLimit(units, 32767); if (!text) return;
    const previous = parts.at(-1);
    if (previous?.kind === "text") parts[parts.length - 1] = { kind: "text", text: previous.text + text };
    else parts.push({ kind: "text", text });
  };
  const padded = (value: number, length: number): string => String(value).padStart(length, "0");
  for (let index = 0; index < section.tokens.length; index++) {
    const token = section.tokens[index]!;
    if (token.kind === "literal") text(token.text);
    else if (token.kind === "slash" || token.kind === "comma") text(token.kind === "slash" ? "/" : ",");
    else if (token.kind === "space" || token.kind === "fill") { units += token.character.length; checkLimit(units, 32767); parts.push(token); }
    else if (token.kind === "dayPeriod") {
      const period = token.abbreviated ? hour < 12 ? "A" : "P" : names.periods[hour < 12 ? 0 : 1];
      text(token.upperCase ? period.toUpperCase() : period.toLowerCase());
    } else if (token.kind === "elapsed") text(padded(Math.floor(seconds / (token.field === "h" ? 3600 : token.field === "m" ? 60 : 1)), token.length));
    else if (token.kind === "decimal") { text(decimal + padded(ticks % unit, precision)); index += plan.precision; }
    else if (token.kind === "date") {
      if (token.field === "y") text(padded(token.length === 2 ? date.year % 100 : date.year, token.length));
      else if (token.field === "d") text(token.length <= 2 ? padded(date.day, token.length) : names.weekdays[token.length === 3 ? "short" : "long"][date.weekday]!);
      else if (token.field === "h") text(padded(twelveHour ? hour % 12 || 12 : hour, token.length));
      else if (token.field === "s") text(padded(second, token.length));
      else if (plan.minutes.has(index)) text(padded(minute, token.length));
      else if (token.length <= 2) text(padded(date.month, token.length));
      else if (token.length === 5) text(names.months.initial[date.month - 1]!);
      else text((plan.dayOfMonth ? names.contextualMonths : names.months)[token.length === 3 ? "short" : "long"][date.month - 1]!);
    }
  }
  return freeze({ kind: "parts", parts });
}
