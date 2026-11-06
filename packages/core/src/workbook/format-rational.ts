import { checkLimit, invalid } from "../document/validation.js";
import { decimalComponents } from "./format-decimal.js";

export interface DisplayFraction { readonly numerator: bigint; readonly denominator: bigint }
const absolute = (value: bigint): bigint => value < 0n ? -value : value;
export function rationalMagnitude(value: number, scale = 0): DisplayFraction {
  if (!Number.isSafeInteger(scale) || Math.abs(scale) > 8192) invalid("Invalid fraction-format decimal scale.");
  const source = decimalComponents(value); if (source.digits === "0") return { numerator: 0n, denominator: 1n };
  const exponent = source.exponent + scale;
  checkLimit(Math.max(source.digits.length + exponent, 1 - exponent), 8192);
  const numerator = BigInt(source.digits);
  return exponent >= 0 ? { numerator: numerator * 10n ** BigInt(exponent), denominator: 1n } : { numerator, denominator: 10n ** BigInt(-exponent) };
}

/** Find the nearest nonnegative fraction with a bounded denominator using
 * continued-fraction convergents. Never scan up to 10^placeholderCount. Exact
 * ties prefer the smaller denominator, then the smaller numerator. */
export function nearestDisplayFraction(source: DisplayFraction, maximumDenominator: bigint): DisplayFraction {
  if (source.numerator < 0n || source.denominator <= 0n || maximumDenominator <= 0n) invalid("Invalid numeric display fraction.");
  for (const value of [source.numerator, source.denominator, maximumDenominator]) checkLimit(value.toString().length, 8192);
  let numerator = source.numerator, denominator = source.denominator;
  let p0 = 0n, q0 = 1n, p1 = 1n, q1 = 0n, steps = 0;
  while (denominator) {
    checkLimit(++steps, 4096);
    const quotient = numerator / denominator, p2 = p0 + quotient * p1, q2 = q0 + quotient * q1;
    if (q2 > maximumDenominator) break;
    [p0, q0, p1, q1] = [p1, q1, p2, q2];
    [numerator, denominator] = [denominator, numerator - quotient * denominator];
  }
  if (!denominator) return { numerator: p1, denominator: q1 };
  const count = (maximumDenominator - q0) / q1;
  const first = { numerator: p0 + count * p1, denominator: q0 + count * q1 }, second = { numerator: p1, denominator: q1 };
  const left = absolute(source.numerator * first.denominator - first.numerator * source.denominator) * second.denominator;
  const right = absolute(source.numerator * second.denominator - second.numerator * source.denominator) * first.denominator;
  return left < right || left === right && (first.denominator < second.denominator || first.denominator === second.denominator && first.numerator < second.numerator) ? first : second;
}
