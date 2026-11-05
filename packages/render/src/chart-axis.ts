import { DocumentError } from "@dagent-office/core";

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
  const span = maximum - minimum, steps = span / unit;
  if (!Number.isSafeInteger(maxSteps) || maxSteps <= 0 || !Number.isFinite(tolerance) || tolerance < 0 || tolerance >= 1
    || ![minimum, maximum, unit, span, steps].every(Number.isFinite) || span <= 0 || unit <= 0 || steps > maxSteps) throw new DocumentError("ResourceLimitExceeded", "Chart axis ticks exceed the finite preview scale or tick budget.");
  const values: number[] = [];
  for (let index = 0; index <= Math.floor(steps + tolerance); index++) {
    const value = minimum + index * unit;
    if (!Number.isFinite(value) || index > 0 && value <= values[index - 1]!) throw new DocumentError("ResourceLimitExceeded", "Chart tick units do not advance at this numeric scale.");
    values.push(value);
  }
  return values;
}
