import { invalid } from "../document/validation.js";

export type WorkbookDateSystem = "1900" | "1904";
export interface WorkbookSerialDate {
  readonly year: number; readonly month: number; readonly day: number;
  /** Sunday is zero, including Excel's historical compatibility weekdays. */
  readonly weekday: number;
  readonly serialDay: number; readonly dayFraction: number;
  readonly compatibility?: "zero-day" | "leap-day";
}

/** Interpret the selected native compatibility calendar, without local time or
 * the machine's timezone. Undefined is an out-of-calendar display value, not a
 * changed cell value or a formula error. Excel's zero day and synthetic leap
 * day cannot be represented by a JavaScript Date and are handled explicitly. */
export function workbookSerialDate(value: number, system: WorkbookDateSystem): WorkbookSerialDate | undefined {
  if (!Number.isFinite(value) || !["1900", "1904"].includes(system)) invalid("A workbook calendar requires a finite serial and a known date system.");
  const lastDay = system === "1900" ? 2_958_465 : 2_957_003;
  if (value < 0 || value >= lastDay + 1) return undefined;
  const serialDay = Math.floor(value), dayFraction = value - serialDay;
  if (system === "1900" && (serialDay === 0 || serialDay === 60)) return Object.freeze({ year: 1900,
    month: serialDay === 0 ? 1 : 2, day: serialDay === 0 ? 0 : 29, weekday: (serialDay + 6) % 7,
    serialDay, dayFraction, compatibility: serialDay === 0 ? "zero-day" : "leap-day" });
  const epoch = system === "1900" ? Date.UTC(1899, 11, 31) : Date.UTC(1904, 0, 1);
  const offset = serialDay - (system === "1900" && serialDay > 60 ? 1 : 0), date = new Date(epoch + offset * 86_400_000);
  return Object.freeze({ year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate(),
    weekday: (serialDay + (system === "1900" ? 6 : 5)) % 7, serialDay, dayFraction });
}
