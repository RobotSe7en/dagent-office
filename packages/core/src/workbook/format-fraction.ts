import { checkLimit, unsupported } from "../document/validation.js";
import { nearestDisplayFraction, rationalMagnitude } from "./format-rational.js";
import type { NumericSectionDisplay, NumberDisplayPart } from "./format-numeric.js";
import type { NumberFormatSection, NumberFormatToken } from "./number-format.js";

type Digit = Extract<NumberFormatToken, { kind: "digit" }>;
const isDigit = (token: NumberFormatToken): token is Digit => token.kind === "digit";

/** One selected fraction section. The denominator is either a literal native
 * integer or a bounded digit field. Quoted digits are never mistaken for a
 * fixed denominator. Values and caches remain unchanged. */
export function formatFractionSection(value: number, section: NumberFormatSection): NumericSectionDisplay {
  const tokens = section.tokens, slashes = tokens.flatMap((token, index) => token.kind === "slash" ? [index] : []), slash = slashes[0];
  if (slashes.length !== 1 || slash === undefined) unsupported("A fraction number format requires one fraction separator.");
  let start = slash, end = slash + 1;
  while (start > 0 && tokens[start - 1]?.kind === "digit") start--;
  const numeratorPattern = tokens.slice(start, slash).filter(isDigit);
  if (!numeratorPattern.length) unsupported("A fraction requires a numerator digit field.");
  const denominatorToken = tokens[end], fixed = denominatorToken?.kind === "integer" ? denominatorToken : undefined;
  if (fixed) end++;
  else while (tokens[end]?.kind === "digit") end++;
  const denominatorPattern = tokens.slice(slash + 1, end).filter(isDigit);
  if (!fixed && !denominatorPattern.length) unsupported("A fraction requires a positive fixed denominator or digit field.");
  const leading = tokens.slice(0, start), integerIndexes = leading.flatMap((token, index) => isDigit(token) ? [index] : []);
  const mixed = integerIndexes.length > 0, first = integerIndexes[0]!, last = integerIndexes.at(-1)!;
  if (mixed && numeratorPattern.some(token => token.value === "0")) unsupported("Mixed fractions with mandatory numerator zeros need additional native layout verification.");
  const integerPattern = mixed ? leading.slice(first, last + 1) : [];
  if (integerPattern.some(token => !isDigit(token))) unsupported("Grouped or interleaved fraction integer fields are not implemented.");
  const prefix = mixed ? leading.slice(0, first) : leading, separator = mixed ? leading.slice(last + 1) : [];
  if (mixed && (!separator.length || separator.some(token => token.kind !== "literal" && token.kind !== "space" && token.kind !== "fill"))) unsupported("Mixed fractions require a literal or measured separator.");
  const source = rationalMagnitude(value, tokens.some(token => token.kind === "percent") ? 2 : 0);
  let whole = mixed ? source.numerator / source.denominator : 0n;
  const remainder = { numerator: source.numerator - whole * source.denominator, denominator: source.denominator };
  const fraction = fixed
    ? { numerator: (2n * remainder.numerator * BigInt(fixed.digits) + remainder.denominator) / (2n * remainder.denominator), denominator: BigInt(fixed.digits) }
    : nearestDisplayFraction(remainder, 10n ** BigInt(denominatorPattern.length) - 1n);
  let numerator = fraction.numerator;
  if (mixed && numerator >= fraction.denominator) { whole += numerator / fraction.denominator; numerator %= fraction.denominator; }
  const parts: NumberDisplayPart[] = [];
  const text = (text: string): void => {
    if (!text) return;
    const previous = parts.at(-1);
    if (previous?.kind === "text") parts[parts.length - 1] = { kind: "text", text: previous.text + text };
    else parts.push({ kind: "text", text });
  };
  const affix = (values: readonly NumberFormatToken[]): void => {
    for (const token of values) {
      if (token.kind === "literal") text(token.text);
      else if (token.kind === "percent") text("%");
      else if (token.kind === "space" || token.kind === "fill") parts.push(token);
      else unsupported("This fraction prefix, separator or suffix is not implemented.");
    }
  };
  const digits = (value: bigint, pattern: readonly Digit[], forceZero = false): void => {
    if (!value && !forceZero) {
      for (const token of pattern) {
        if (token.value === "0") text("0");
        else if (token.value === "?") parts.push({ kind: "space", character: "0" });
      }
    } else {
      const firstRequired = pattern.findIndex(token => token.value === "0");
      const rendered = value.toString().padStart(firstRequired < 0 ? 0 : pattern.length - firstRequired, "0");
      for (let index = 0; index < pattern.length - rendered.length; index++) if (pattern[index]!.value === "?") parts.push({ kind: "space", character: "0" });
      text(rendered);
    }
  };
  affix(prefix);
  if (mixed) {
    digits(whole, integerPattern.filter(isDigit), whole === 0n && numerator === 0n);
    affix(separator);
  }
  const hidden = mixed && numerator === 0n && !numeratorPattern.some(token => token.value === "0");
  if (hidden) {
    for (const token of numeratorPattern) if (token.value === "?") parts.push({ kind: "space", character: "0" });
    parts.push({ kind: "space", character: "/" });
    for (const token of denominatorPattern) if (token.value !== "#") parts.push({ kind: "space", character: "0" });
  } else {
    digits(numerator, numeratorPattern, true); text("/");
    if (fixed) text(fixed.digits); else digits(fraction.denominator, denominatorPattern, true);
  }
  affix(tokens.slice(end));
  checkLimit(parts.reduce((sum, part) => sum + (part.kind === "text" ? part.text.length : part.character.length), 0), 32767);
  return { parts, zero: whole === 0n && numerator === 0n };
}
