import { DocumentError } from "@dagent-office/core";
import { chartValueFraction } from "./chart-axis.js";

export const MAX_CHART_PAINT_EXTENT = 1_000_000;
export interface ChartValuePoint { readonly x: number; readonly value: number }
export interface ChartClippedPoint { readonly x: number; readonly fraction: number }

/** Divide intervals without first overflowing either difference. Source
 * endpoints stay doubles; only a bounded segment's intersection is projected. */
function spanRatio(low: number, high: number, from: number, to: number): number {
  const numerator = high - low, denominator = to - from;
  if (Number.isFinite(numerator) && Number.isFinite(denominator)) return numerator / denominator;
  const scale = Math.max(Math.abs(low), Math.abs(high), Math.abs(from), Math.abs(to));
  return (high / scale - low / scale) / (to / scale - from / scale);
}

/** Clip in the value domain before creating SVG coordinates. Intersections
 * interpolate original category positions; clamping each value independently
 * would change the slope. Fractional clipping bounds also retain subnormal
 * axes whose expanded endpoints cannot themselves be represented as doubles. */
export function clipChartValueSegment(first: ChartValuePoint, last: ChartValuePoint, minimum: number, maximum: number, lowFraction: number, highFraction: number): readonly [ChartClippedPoint, ChartClippedPoint] | undefined {
  if (![first.x, first.value, last.x, last.value, minimum, maximum, lowFraction, highFraction].every(Number.isFinite) || maximum <= minimum || highFraction <= lowFraction)
    throw new DocumentError("InvalidOperation", "Chart clipping requires finite endpoints and ordered intervals.");
  const from = chartValueFraction(first.value, minimum, maximum), to = chartValueFraction(last.value, minimum, maximum);
  if (from < lowFraction && to < lowFraction || from > highFraction && to > highFraction) return;
  const endpoint = (point: ChartValuePoint, fraction: number): ChartClippedPoint => {
    if (fraction >= lowFraction && fraction <= highFraction) return Object.freeze({ x: point.x, fraction });
    const boundary = fraction < lowFraction ? lowFraction : highFraction;
    const t = Math.max(0, Math.min(1, chartValueFraction(minimum, first.value, last.value) + boundary * spanRatio(minimum, maximum, first.value, last.value)));
    const x = first.x + (last.x - first.x) * t;
    if (!Number.isFinite(x)) throw new DocumentError("ResourceLimitExceeded", "A chart intersection exceeds finite category coordinates.");
    return Object.freeze({ x, fraction: boundary });
  };
  return Object.freeze([endpoint(first, from), endpoint(last, to)]);
}

export function chartStrokeOutset(stroke: Readonly<Record<string, string | number>>): number {
  const width = Number(stroke["stroke-width"]), miter = Number(stroke["stroke-miterlimit"] ?? 1);
  if (![width, miter].every(value => Number.isFinite(value) && value >= 0)) throw new DocumentError("InvalidDocument", "Chart stroke bounds must be finite and nonnegative.");
  const outset = stroke.stroke === "none" ? 0 : width / 2 * Math.max(stroke["stroke-linecap"] === "square" ? Math.SQRT2 : 1, stroke["stroke-linejoin"] === "miter" ? miter : 1);
  if (!Number.isFinite(outset) || outset > MAX_CHART_PAINT_EXTENT) throw new DocumentError("ResourceLimitExceeded", "A chart stroke exceeds its finite paint extent budget.");
  return outset;
}
