import { checkLimit, invalid, unsupported } from "../document/validation.js";
import { decimalComponents, decimalMagnitude } from "./format-decimal.js";
import { formatDecimalSection } from "./format-numeric.js";
import type { NumericSectionDisplay, NumberDisplayPart, NumberSeparators } from "./format-numeric.js";
import type { NumberFormatSection } from "./number-format.js";

/** Scientific display shifts decimal digits, not the stored floating number.
 * Optional leading integer places select an engineering exponent interval. */
export function formatScientificSection(value: number, section: NumberFormatSection, separators: NumberSeparators): NumericSectionDisplay {
  if (!Number.isFinite(value)) invalid("Scientific formatting requires a finite stored scalar.");
  const indexes = section.tokens.flatMap((token, index) => token.kind === "exponent" ? [index] : []), index = indexes[0];
  if (indexes.length !== 1 || index === undefined) unsupported("A scientific number format requires one exponent field.");
  const token = section.tokens[index]!; if (token.kind !== "exponent") invalid("Missing scientific exponent marker.");
  let end = index + 1;
  while (section.tokens[end]?.kind === "digit") end++;
  const exponentDigits = section.tokens.slice(index + 1, end);
  if (!exponentDigits.length || exponentDigits.some(token => token.kind !== "digit" || token.value !== "0")) unsupported("Scientific exponent fields require mandatory decimal digits.");
  const mantissa = section.tokens.slice(0, index), point = mantissa.findIndex(token => token.kind === "decimal");
  const integer = mantissa.slice(0, point < 0 ? mantissa.length : point).filter(token => token.kind === "digit");
  const fraction = point < 0 ? 0 : mantissa.slice(point + 1).filter(token => token.kind === "digit").length;
  if (!integer.length || !integer.some(token => token.kind === "digit" && token.value === "0")) unsupported("A scientific mantissa requires a mandatory integer digit.");
  const engineering = integer.some(token => token.kind === "digit" && token.value !== "0"), interval = engineering ? integer.length : 1;
  const lastDigit = mantissa.reduce((last, token, index) => token.kind === "digit" ? index : last, -1); let commas = 0;
  while (mantissa[lastDigit + 1 + commas]?.kind === "comma") commas++;
  const percent = section.tokens.some(token => token.kind === "percent"), mantissaPercent = mantissa.some(token => token.kind === "percent");
  const scale = (percent ? 2 : 0) - 3 * commas;
  let exponent = 0;
  if (value !== 0) {
    const source = decimalComponents(value), leading = source.digits.length + source.exponent - 1 + scale;
    exponent = engineering ? Math.floor(leading / interval) * interval : leading - integer.length + 1;
    const rounded = decimalMagnitude(value, fraction, scale - exponent);
    if (rounded.integer.length > integer.length) exponent += interval;
  }
  const displayed = formatDecimalSection(value, { tokens: mantissa }, separators, -exponent + (percent ? 2 : 0) - (mantissaPercent ? 2 : 0));
  const parts: NumberDisplayPart[] = [...displayed.parts];
  const text = (value: string): void => {
    const previous = parts.at(-1);
    if (previous?.kind === "text") parts[parts.length - 1] = { kind: "text", text: previous.text + value };
    else parts.push({ kind: "text", text: value });
  };
  text((token.upperCase ? "E" : "e") + (exponent < 0 ? "-" : token.positiveSign ? "+" : "") + Math.abs(exponent).toString().padStart(exponentDigits.length, "0"));
  for (const suffix of section.tokens.slice(end)) {
    if (suffix.kind === "literal") text(suffix.text);
    else if (suffix.kind === "percent") text("%");
    else if (suffix.kind === "space" || suffix.kind === "fill") parts.push(suffix);
    else unsupported("The scientific number-format suffix is not implemented.");
  }
  checkLimit(parts.reduce((sum, part) => sum + (part.kind === "text" ? part.text.length : part.character.length), 0), 32767);
  return { parts, zero: displayed.zero };
}
