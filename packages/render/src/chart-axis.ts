import { DocumentError } from "@dagent-office/core";

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
