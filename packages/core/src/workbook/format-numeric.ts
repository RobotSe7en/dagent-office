import { checkLimit, invalid, unsupported } from "../document/validation.js";
import { decimalMagnitude } from "./format-decimal.js";
import type { NumberFormatSection, NumberFormatToken } from "./number-format.js";

export type NumberDisplayPart = {
  readonly kind: "text";
  readonly text: string;
  /** Original text-value offsets for an @ field. Literals and numeric/date
   * displays have no source range, even if their characters happen to match. */
  readonly source?: { readonly from: number; readonly to: number };
} | { readonly kind: "space" | "fill"; readonly character: string };
export interface NumberSeparators { readonly decimal: string; readonly group: string; readonly primaryGroupSize: number; readonly secondaryGroupSize: number }
export interface NumericSectionDisplay {
  readonly parts: readonly NumberDisplayPart[];
  /** True only when the displayed magnitude rounds to zero. Sign/section
   * selection belongs to the caller, rather than mutating the stored scalar. */
  readonly zero: boolean;
}

/** Display one ordinary decimal section's magnitude. Colors, conditions and
 * locale identifiers remain on its compiled section. Space/fill characters are
 * measured by the renderer; they are never flattened to an arbitrary space. */
export function formatDecimalSection(value: number, section: NumberFormatSection, separators: NumberSeparators, powerAdjustment = 0): NumericSectionDisplay {
  if (!Number.isFinite(value) || !separators.decimal || separators.decimal.length > 8 || !separators.group || separators.group.length > 8
    || ![separators.primaryGroupSize, separators.secondaryGroupSize].every(size => Number.isSafeInteger(size) && size > 0 && size <= 9)
    || !Number.isSafeInteger(powerAdjustment) || Math.abs(powerAdjustment) > 8192) invalid("Invalid scalar or numeric separator configuration.");
  const tokens = section.tokens, digits = tokens.flatMap((token, index) => token.kind === "digit" ? [index] : []);
  if (!digits.length || tokens.some(token => ["date", "elapsed", "dayPeriod", "exponent", "slash", "text", "general"].includes(token.kind))) unsupported("This section requires a non-decimal number-format renderer.");
  const decimals = tokens.flatMap((token, index) => token.kind === "decimal" ? [index] : []);
  if (decimals.length > 1) unsupported("Multiple number-format decimal fields are not implemented.");
  const decimal = decimals[0], start = Math.min(digits[0]!, decimal ?? Infinity), lastDigit = digits.at(-1)!;
  if (decimal !== undefined && decimal > lastDigit + 1) unsupported("A separated trailing decimal marker is not implemented.");
  const end = Math.max(lastDigit + 1, decimal === undefined ? 0 : decimal + 1);
  const integerTokens = tokens.slice(start, decimal ?? end), fractionTokens = decimal === undefined ? [] : tokens.slice(decimal + 1, end);
  if (integerTokens.some(token => token.kind !== "digit" && token.kind !== "comma") || fractionTokens.some(token => token.kind !== "digit")
    || integerTokens[0]?.kind === "comma" || integerTokens.at(-1)?.kind === "comma") unsupported("Interleaved number-format fields require a different numeric pattern.");
  const integerPattern = integerTokens.flatMap(token => token.kind === "digit" ? [token.value] : []), fractionPattern = fractionTokens.map(token => (token as Extract<NumberFormatToken, { kind: "digit" }>).value);
  let after = end, commas = 0;
  while (tokens[after]?.kind === "comma") { commas++; after++; }
  const percent = tokens.some(token => token.kind === "percent"), scale = (percent ? 2 : 0) - commas * 3 + powerAdjustment;
  const rounded = decimalMagnitude(value, fractionPattern.length, scale), parts: NumberDisplayPart[] = [];
  const text = (value: string): void => {
    if (!value) return;
    const previous = parts.at(-1);
    if (previous?.kind === "text") parts[parts.length - 1] = { kind: "text", text: previous.text + value };
    else parts.push({ kind: "text", text: value });
  };
  const affix = (values: readonly NumberFormatToken[]): void => {
    for (const token of values) {
      if (token.kind === "literal") text(token.text);
      else if (token.kind === "percent") text("%");
      else if (token.kind === "space" || token.kind === "fill") parts.push(token);
      else unsupported("This numeric prefix or suffix is not implemented.");
    }
  };
  affix(tokens.slice(0, start));
  const firstRequired = integerPattern.indexOf("0"), minimumDigits = firstRequired < 0 ? 0 : integerPattern.length - firstRequired;
  const grouping = integerTokens.some(token => token.kind === "comma");
  const group = (integer: string): string => {
    if (!grouping || !integer) return integer;
    const groups: string[] = []; let end = integer.length, size = separators.primaryGroupSize;
    while (end > 0) { const start = Math.max(0, end - size); groups.push(integer.slice(start, end)); end = start; size = separators.secondaryGroupSize; }
    return groups.reverse().join(separators.group);
  };
  if (rounded.integer === "0") {
    let remaining = integerPattern.filter(value => value === "0").length, emitted = 0;
    for (const placeholder of integerPattern) {
      if (placeholder === "0") {
        if (grouping && emitted && (remaining === separators.primaryGroupSize || remaining > separators.primaryGroupSize && (remaining - separators.primaryGroupSize) % separators.secondaryGroupSize === 0)) text(separators.group);
        text("0"); remaining--; emitted++;
      } else if (placeholder === "?") parts.push({ kind: "space", character: "0" });
    }
  } else {
    const integer = rounded.integer.padStart(minimumDigits, "0");
    for (let index = 0; index < integerPattern.length - integer.length; index++) if (integerPattern[index] === "?") parts.push({ kind: "space", character: "0" });
    text(group(integer));
  }
  let lastRequired = -1, lastSignificant = -1;
  for (let index = 0; index < fractionPattern.length; index++) {
    if (fractionPattern[index] === "0") lastRequired = index;
    if (rounded.fraction[index] !== "0") lastSignificant = index;
  }
  if (lastRequired >= 0 || lastSignificant >= 0 || fractionPattern.includes("?")) text(separators.decimal);
  for (let index = 0; index < fractionPattern.length; index++) {
    if (fractionPattern[index] === "?" && index > lastSignificant) parts.push({ kind: "space", character: "0" });
    else if (index <= Math.max(lastRequired, lastSignificant)) text(rounded.fraction[index]!);
  }
  affix(tokens.slice(after));
  checkLimit(parts.reduce((sum, part) => sum + (part.kind === "text" ? part.text.length : part.character.length), 0), 32767);
  return { parts, zero: rounded.integer === "0" && !/[1-9]/.test(rounded.fraction) };
}
