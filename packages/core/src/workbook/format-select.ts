import { checkLimit, invalid, unsupported } from "../document/validation.js";
import type { NumberDisplayPart } from "./format-numeric.js";
import type { NumberFormatCondition, NumberFormatSection } from "./number-format.js";

export type NumberFormatSelection =
  | { readonly kind: "number"; readonly index: number; readonly section: NumberFormatSection; readonly automaticNegative: boolean }
  | { readonly kind: "text"; readonly index: number; readonly section: NumberFormatSection }
  | { readonly kind: "general" }
  | { readonly kind: "plainText" }
  | { readonly kind: "unmatchedCondition" };

function matches(value: number, condition: NumberFormatCondition): boolean {
  switch (condition.operator) {
    case "<": return value < condition.value;
    case "<=": return value <= condition.value;
    case ">": return value > condition.value;
    case ">=": return value >= condition.value;
    case "=": return value === condition.value;
    case "<>": return value !== condition.value;
  }
}

/** Select against the original scalar, before display rounding. A final @ field
 * is a text section even when fewer than four sections were supplied. Failure
 * to match explicit conditions stays distinct from an implicit General format. */
export function selectNumberFormat(value: number | string, sections: readonly NumberFormatSection[]): NumberFormatSelection {
  if (!sections.length || sections.length > 4 || typeof value === "number" && !Number.isFinite(value)) invalid("Invalid number-format value or section count.");
  const last = sections.length - 1, hasText = sections[last]!.tokens.some(token => token.kind === "text");
  const textIndex = sections.length === 4 || hasText ? last : undefined, count = textIndex ?? sections.length;
  if (sections.slice(0, count).some(section => section.tokens.some(token => token.kind === "text"))) unsupported("A number format can only have a text field in its last section.");
  if (textIndex !== undefined && sections[textIndex]!.condition) unsupported("A text section cannot carry a numeric condition.");
  if (typeof value === "string") return textIndex === undefined ? { kind: "plainText" } : { kind: "text", index: textIndex, section: sections[textIndex]! };
  if (!count) return { kind: "general" };
  const selected = (index: number, automaticNegative: boolean): NumberFormatSelection => ({ kind: "number", index, section: sections[index]!, automaticNegative });
  if (sections[1]?.condition && !sections[0]?.condition) unsupported("A condition in only the second numeric section needs additional native verification.");
  const first = sections[0]!;
  if (first.condition) {
    if (count === 3 && !sections[1]?.condition) unsupported("Three numeric sections with only one explicit condition need additional native verification.");
    if (matches(value, first.condition)) return selected(0, false);
    const second = count > 1 ? sections[1] : undefined;
    if (second && (!second.condition || matches(value, second.condition))) return selected(1, false);
    return count > 2 ? selected(2, false) : { kind: "unmatchedCondition" };
  }
  return selected(value < 0 && count >= 2 ? 1 : value === 0 && count >= 3 ? 2 : 0, count === 1 && value < 0);
}

/** Render text/literal-only sections without parsing the input as another value
 * type. Expansion is checked before concatenation, including repeated @ fields. */
export function formatLiteralSection(value: string | undefined, section: NumberFormatSection, retainTextSource = false): readonly NumberDisplayPart[] {
  const parts: NumberDisplayPart[] = []; let units = 0;
  const text = (text: string): void => {
    units += text.length; checkLimit(units, 32767); if (!text) return;
    const previous = parts.at(-1);
    if (previous?.kind === "text" && !previous.source) parts[parts.length - 1] = { kind: "text", text: previous.text + text };
    else parts.push({ kind: "text", text });
  };
  for (const token of section.tokens) {
    if (token.kind === "literal") text(token.text);
    else if (token.kind === "text" && value !== undefined) {
      if (retainTextSource) { units += value.length; checkLimit(units, 32767); parts.push({ kind: "text", text: value, source: { from: 0, to: value.length } }); }
      else text(value);
    }
    else if (token.kind === "space" || token.kind === "fill") { units += token.character.length; checkLimit(units, 32767); parts.push(token); }
    else unsupported("This text or literal number-format section contains a nonliteral field.");
  }
  return parts;
}
