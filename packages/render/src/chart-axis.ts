import { DocumentError } from "@dagent-office/core";
import type { ChartAppearance, DeepReadonly } from "@dagent-office/core";
import { property } from "./presentation/appearance.js";

export function chartValueScale(chart: DeepReadonly<ChartAppearance>, series: readonly { readonly hidden?: boolean; readonly values: readonly (number | null)[] }[]): {
  readonly minimum: number; readonly maximum: number; readonly unit: number;
} {
  let low = 0, high = 0;
  for (const item of series) if (!item.hidden) for (const value of item.values) if (value !== null) { low = Math.min(low, value); high = Math.max(high, value); }
  if (low === 0 && high === 0) high = 1;
  const minimum = chart.minimum ?? low, largest = chart.maximum ?? high;
  const span = largest - minimum;
  const rough = span > 0 ? Math.max(Number.MIN_VALUE, Number.isFinite(span) ? span / 8 : largest / 8 - minimum / 8) : 1;
  const power = Math.max(Number.MIN_VALUE, 10 ** Math.floor(Math.log10(rough))), factor = rough / power;
  const unit = chart.majorUnit ?? (factor <= 1 ? 1 : factor <= 2 ? 2 : factor <= 5 ? 5 : 10) * power;
  const automatic = Math.min(Number.MAX_VALUE, (Math.floor(largest / unit) + 1) * unit);
  return { minimum, maximum: chart.maximum ?? Math.max(automatic, Math.min(Number.MAX_VALUE, minimum + unit)), unit };
}

/** Normalize finite values without requiring the full interval or numerator
 * to fit in a double. Preserve the direct path's precision for ordinary data. */
export function chartValueFraction(value: number, minimum: number, maximum: number, reversed = false): number {
  const span = maximum - minimum, offset = reversed ? maximum - value : value - minimum;
  if (Number.isFinite(span) && Number.isFinite(offset)) return offset / span;
  const scale = Math.max(Math.abs(value), Math.abs(minimum), Math.abs(maximum));
  return (reversed ? maximum / scale - value / scale : value / scale - minimum / scale) / (maximum / scale - minimum / scale);
}

/** Fractions run left-to-right / top-to-bottom. Native axPos remains source
 * data but never participates in Office placement (MS-OI29500 §2.1.1433).
 * crossesAt is measured on the perpendicular axis, not its parent axis. */
export function chartAxisPlacement(chart: DeepReadonly<ChartAppearance>, category: CSSStyleDeclaration, value: CSSStyleDeclaration,
  minimum: number, maximum: number, count: number): {
    readonly x: number; readonly y: number; readonly categoryTop: boolean; readonly valueRight: boolean;
    readonly categoryLabel: number; readonly valueLabel: number;
    readonly categoryOutside: number; readonly valueOutside: number;
  } {
  const between = property(value, "--office-axis-cross-between", "between");
  if (!["between", "midCat"].includes(between) || !Number.isSafeInteger(count) || count < 1 || ![minimum, maximum].every(Number.isFinite) || maximum <= minimum)
    throw new DocumentError("InvalidDocument", "Chart axis intervals are invalid.");
  const crossing = (style: CSSStyleDeclaration, low: number, high: number, integral = false): number => {
    const raw = property(style, "--office-axis-crosses-at"), kind = property(style, "--office-axis-crosses", "autoZero"), position = property(style, "--office-axis-position");
    if (!["min", "max", "autoZero"].includes(kind) || position && !["b", "t", "l", "r"].includes(position)) throw new DocumentError("InvalidDocument", "Chart axis crossing options are invalid.");
    let at = kind === "max" ? high : kind === "min" ? low : 0;
    if (raw) {
      at = Number(raw);
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?$/.test(raw) || !Number.isFinite(at) || integral && !Number.isInteger(at)) throw new DocumentError("InvalidDocument", "Chart axis crossing numbers are invalid.");
    }
    return Math.max(low, Math.min(high, at));
  };
  const slots = between === "between" ? count : Math.max(1, count - 1);
  const categoryAt = crossing(category, minimum, maximum), valueAt = crossing(value, 1, slots + 1, true);
  const vertical = (at: number): number => chart.reverseValues ? at : 1 - at;
  const horizontal = (at: number): number => chart.reverseCategories ? 1 - at : at;
  const y = vertical(chartValueFraction(categoryAt, minimum, maximum)), x = horizontal((valueAt - 1) / slots);
  const labels = (style: CSSStyleDeclaration, at: number, project: (at: number) => number): number => {
    const option = property(style, "--office-axis-label-position", "nextTo");
    if (!["high", "low", "nextTo", "none"].includes(option)) throw new DocumentError("InvalidDocument", "Chart tick-label placement is invalid.");
    return option === "high" ? project(1) : option === "low" ? project(0) : at;
  };
  const categoryLabel = labels(category, y, vertical), valueLabel = labels(value, x, horizontal);
  // Interior nextTo labels use the low side; reversed perpendicular axes
  // reverse that side. At the high boundary they face outside the plot.
  const categoryTop = categoryLabel === vertical(1) ? !chart.reverseValues : !!chart.reverseValues;
  const valueRight = valueLabel === horizontal(1) ? !chart.reverseCategories : !!chart.reverseCategories;
  const categoryOutside = (y === vertical(1) ? !chart.reverseValues : !!chart.reverseValues) ? -1 : 1;
  const valueOutside = (x === horizontal(1) ? !chart.reverseCategories : !!chart.reverseCategories) ? 1 : -1;
  return { x, y, categoryLabel, valueLabel, categoryTop, valueRight, categoryOutside, valueOutside };
}

/** The same category interval positions the current points, grid and ticks.
 * A singleton keeps one finite interval even when native crossing is midCat. */
export function chartCategoryAxis(count: number, left: number, width: number, between: boolean, reverse: boolean): {
  readonly interval: number; readonly tickCount: number; readonly tick: (index: number) => number; readonly position: (index: number) => number;
} {
  if (!Number.isSafeInteger(count) || count < 1 || count > 0xffff_ffff || ![left, width].every(Number.isFinite) || width <= 0) throw new DocumentError("ResourceLimitExceeded", "A chart category axis exceeds its coordinate or point budget.");
  const interval = width / (between ? count : Math.max(1, count - 1));
  const position = (index: number): number => { const local = (index + (between ? .5 : 0)) * interval; return left + (reverse ? width - local : local); };
  return Object.freeze({ interval, tickCount: between ? count + 1 : Math.max(2, count), tick: (index: number) => left + index * interval, position });
}

/** Index bounded ticks rather than repeatedly adding a floating-point unit.
 * At large magnitudes a positive unit may not advance the value at all; reject
 * that scale even when axes/ticks are hidden and no node budget would fire. */
export function chartAxisTicks(minimum: number, maximum: number, unit: number, maxSteps: number, tolerance = 0): readonly number[] {
  const span = maximum - minimum, steps = Number.isFinite(span) ? span / unit : maximum / unit - minimum / unit;
  if (!Number.isSafeInteger(maxSteps) || maxSteps <= 0 || !Number.isFinite(tolerance) || tolerance < 0 || tolerance >= 1
    || ![minimum, maximum, unit, steps].every(Number.isFinite) || maximum <= minimum || unit <= 0 || steps > maxSteps) throw new DocumentError("ResourceLimitExceeded", "Chart axis ticks exceed the finite preview scale or tick budget.");
  const values: number[] = [];
  for (let index = 0; index <= Math.floor(steps + tolerance); index++) {
    const offset = index * unit, value = Number.isFinite(offset) ? minimum + offset : (minimum / unit + index) * unit;
    if (!Number.isFinite(value) || index > 0 && value <= values[index - 1]!) throw new DocumentError("ResourceLimitExceeded", "Chart tick units do not advance at this numeric scale.");
    values.push(value);
  }
  return values;
}
