import { checkLimit, freeze, invalid, unsupported } from "../document/validation.js";

export type NumberFormatToken =
  | { readonly kind: "literal"; readonly text: string }
  | { readonly kind: "digit"; readonly value: "0" | "#" | "?" }
  | { readonly kind: "integer"; readonly digits: string }
  | { readonly kind: "decimal" | "comma" | "percent" | "slash" | "text" | "general" }
  | { readonly kind: "exponent"; readonly positiveSign: boolean; readonly upperCase: boolean }
  | { readonly kind: "date"; readonly field: "y" | "m" | "d" | "h" | "s"; readonly length: number }
  | { readonly kind: "elapsed"; readonly field: "h" | "m" | "s"; readonly length: number }
  | { readonly kind: "dayPeriod"; readonly abbreviated: boolean; readonly upperCase: boolean }
  | { readonly kind: "space" | "fill"; readonly character: string };
export interface NumberFormatCondition { readonly operator: "<" | "<=" | ">" | ">=" | "=" | "<>"; readonly value: number }
export type NumberFormatColor = { readonly kind: "named"; readonly name: "black" | "blue" | "cyan" | "green" | "magenta" | "red" | "white" | "yellow" } | { readonly kind: "indexed"; readonly index: number };
export interface NumberFormatSection {
  readonly tokens: readonly NumberFormatToken[];
  readonly condition?: NumberFormatCondition;
  readonly color?: NumberFormatColor;
  /** Native LCID/format flags, not the host's current UI locale. */
  readonly locale?: number;
}
interface PendingSection { tokens: NumberFormatToken[]; condition?: NumberFormatCondition; color?: NumberFormatColor; locale?: number }
const names = new Set(["black", "blue", "cyan", "green", "magenta", "red", "white", "yellow"]);
const numeric = /^([<>]=?|=|<>)([+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?)$/;

/** Tokenize the original code before interpreting sections or punctuation.
 * Quoted/escaped semicolons, percent signs and date letters remain literals;
 * fill/spacing retain their character for later measured layout. No values,
 * locale choices or date-system assumptions enter this source compiler. */
export function compileNumberFormat(code: string): readonly NumberFormatSection[] {
  if (!code || /[\u0000-\u0008\u000b\u000c\u000e-\u001f]/.test(code)) invalid("A number format requires a nonempty source without invalid control characters.");
  checkLimit(code.length, 2048);
  for (const character of code) { const point = character.codePointAt(0)!; if (point >= 0xd800 && point <= 0xdfff) invalid("Number format characters must be well-formed Unicode."); }
  const sections: PendingSection[] = [{ tokens: [] }]; let section = sections[0]!, cursor = 0;
  const literal = (text: string): void => {
    const last = section.tokens.at(-1);
    if (last?.kind === "literal") section.tokens[section.tokens.length - 1] = { kind: "literal", text: last.text + text };
    else if (text) section.tokens.push({ kind: "literal", text });
  };
  const character = (): string => {
    if (cursor >= code.length) invalid("A number-format escape, spacing or fill marker requires a following character.");
    const point = code.codePointAt(cursor)!;
    const result = String.fromCodePoint(point); cursor += result.length; return result;
  };
  while (cursor < code.length) {
    const start = cursor, value = character();
    if (value === ";") {
      if (sections.length === 4) invalid("A number format cannot contain more than four sections.");
      section = { tokens: [] }; sections.push(section);
    } else if (value === '"') {
      let text = "", closed = false;
      while (cursor < code.length) { const next = character(); if (next === '"') { closed = true; break; } text += next; }
      if (!closed) invalid("A number-format literal is missing its closing quote."); literal(text);
    } else if (value === "\\") literal(character());
    else if (value === "_" || value === "*") section.tokens.push({ kind: value === "_" ? "space" : "fill", character: character() });
    else if (value === "[") {
      const end = code.indexOf("]", cursor); if (end < 0) invalid("A number-format bracket is not closed.");
      const content = code.slice(cursor, end); cursor = end + 1;
      const lowered = content.toLowerCase(), condition = numeric.exec(content), indexed = /^color([1-9]\d*)$/i.exec(content), elapsed = /^([hms])\1?$/i.exec(content);
      if (names.has(lowered) || indexed) {
        if (section.color) invalid("A number-format section has competing color declarations.");
        if (indexed) {
          const index = Number(indexed[1]); if (index > 56) invalid("A number-format palette index must lie in 1–56.");
          section.color = { kind: "indexed", index };
        } else section.color = { kind: "named", name: lowered as Extract<NumberFormatColor, { kind: "named" }>["name"] };
      } else if (condition) {
        if (section.condition || sections.length > 2 || !Number.isFinite(Number(condition[2]))) invalid("A number format has an invalid or repeated condition.");
        section.condition = { operator: condition[1] as NumberFormatCondition["operator"], value: Number(condition[2]) };
      } else if (elapsed) section.tokens.push({ kind: "elapsed", field: lowered[0] as "h" | "m" | "s", length: lowered.length });
      else if (content.startsWith("$")) {
        const currency = /^\$(.*?)(?:-([0-9a-fA-F]{1,8}))?$/.exec(content);
        if (!currency) unsupported("This native number-format currency or locale declaration is not supported.");
        if (currency[2]) {
          if (section.locale !== undefined) invalid("A number-format section has competing locale declarations.");
          section.locale = parseInt(currency[2], 16);
        }
        literal(currency[1]!);
      } else unsupported("This number-format calendar, digit system or bracket extension is not implemented.");
    } else if ("0#?".includes(value)) section.tokens.push({ kind: "digit", value: value as "0" | "#" | "?" });
    else if (/^[1-9]$/.test(value)) {
      while (code[cursor] && /^\d$/.test(code[cursor]!)) cursor++;
      section.tokens.push({ kind: "integer", digits: code.slice(start, cursor) });
    }
    else if (value === ".") section.tokens.push({ kind: "decimal" });
    else if (value === ",") section.tokens.push({ kind: "comma" });
    else if (value === "%") section.tokens.push({ kind: "percent" });
    else if (value === "@") section.tokens.push({ kind: "text" });
    else if (/^general/i.test(code.slice(start))) { cursor = start + 7; section.tokens.push({ kind: "general" }); }
    else if (/^am\/pm/i.test(code.slice(start)) || /^a\/p/i.test(code.slice(start))) {
      const abbreviated = !/^am\/pm/i.test(code.slice(start)); cursor = start + (abbreviated ? 3 : 5);
      section.tokens.push({ kind: "dayPeriod", abbreviated, upperCase: value === "A" });
    } else if (/^[eE][+-](?=[0#?])/.test(code.slice(start))) {
      section.tokens.push({ kind: "exponent", positiveSign: code[cursor] === "+", upperCase: value === "E" }); cursor++;
    } else if (/^[ymdhs]$/i.test(value)) {
      while (code[cursor]?.toLowerCase() === value.toLowerCase()) cursor++;
      section.tokens.push({ kind: "date", field: value.toLowerCase() as "y" | "m" | "d" | "h" | "s", length: cursor - start });
    } else if (value === "/") section.tokens.push({ kind: "slash" });
    else if (/^[a-z]$/i.test(value) || value === "[" || value === "]") unsupported("Unquoted number-format letters or calendar fields are outside the implemented grammar.");
    else literal(value);
  }
  return freeze(sections);
}
