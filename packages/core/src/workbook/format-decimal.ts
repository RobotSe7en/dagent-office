import { checkLimit, invalid } from "../document/validation.js";

/** The shortest decimal representation of the stored magnitude, with all
 * decimal-point positioning expressed as a power of ten. */
export function decimalComponents(value: number): { readonly digits: string; readonly exponent: number } {
  if (!Number.isFinite(value)) invalid("Number formatting requires a finite stored scalar.");
  const [mantissa, exponentText = "0"] = Math.abs(value).toString().split("e"), [integer, fraction = ""] = mantissa!.split(".");
  return { digits: (integer! + fraction).replace(/^0+/, "") || "0", exponent: Number(exponentText) - fraction.length };
}

/** Decimal display rounding over a finite stored scalar. Scaling is an integer
 * power of ten, so percent/comma formatting cannot overflow an otherwise valid
 * value through an intermediate floating-point multiplication. This does not
 * change the source number or implement formula ROUND. */
export function decimalMagnitude(value: number, fractionPlaces: number, scale = 0): { readonly integer: string; readonly fraction: string } {
  if (!Number.isFinite(value) || !Number.isSafeInteger(fractionPlaces) || fractionPlaces < 0 || fractionPlaces > 2048
    || !Number.isSafeInteger(scale) || Math.abs(scale) > 8192) invalid("Invalid decimal number-format precision or scale.");
  const source = decimalComponents(value), digits = source.digits, exponent = source.exponent + scale + fractionPlaces;
  let units: string;
  if (digits === "0") units = "0";
  else if (exponent >= 0) { checkLimit(digits.length + exponent, 8192); units = digits + "0".repeat(exponent); }
  else if (-exponent > digits.length) units = "0";
  else {
    const divisor = 10n ** BigInt(-exponent);
    units = ((BigInt(digits) + divisor / 2n) / divisor).toString();
  }
  checkLimit(Math.max(units.length, fractionPlaces + 1), 8192);
  const padded = units.padStart(fractionPlaces + 1, "0");
  return fractionPlaces ? { integer: padded.slice(0, -fractionPlaces), fraction: padded.slice(-fractionPlaces) } : { integer: padded, fraction: "" };
}
