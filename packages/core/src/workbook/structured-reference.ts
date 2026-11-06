import { invalid, unsupported } from "../document/validation.js";

export type TableSelector = "all" | "data" | "headers" | "totals" | "thisRow";
export interface FormulaStructuredReference {
  readonly kind: "structured";
  readonly table?: string;
  readonly selectors: readonly TableSelector[];
  /** Empty means all columns; two entries delimit an inclusive column range. */
  readonly columns: readonly string[];
}
const keywords: Readonly<Record<string, TableSelector>> = {
  "#all": "all", "#data": "data", "#headers": "headers", "#totals": "totals", "#this row": "thisRow",
};
const labels: Readonly<Record<TableSelector, string>> = { all: "#All", data: "#Data", headers: "#Headers", totals: "#Totals", thisRow: "#This Row" };

/** Consume one balanced native table specifier; apostrophe escapes delimiters. */
export function readStructuredToken(source: string, start: number): { text: string; end: number } {
  let depth = 0;
  for (let index = start; index < source.length; index++) {
    const character = source[index]!;
    if (character === "'") {
      if (!"'[]#@".includes(source[index + 1] ?? "") || index + 1 >= source.length) invalid("Invalid table column escape.");
      index++; continue;
    }
    if (character === "[") { if (++depth > 2) unsupported("Nested table specifier exceeds the supported grammar."); }
    if (character === "]" && --depth === 0) return { text: source.slice(start, index + 1), end: index + 1 };
  }
  invalid("Unclosed table reference.");
}

export function parseStructuredReference(text: string, table?: string): FormulaStructuredReference {
  let cursor = 0;
  const selectors: TableSelector[] = [], columns: string[] = [];
  let columnRange = false;
  const skip = (): void => { while (text[cursor] === " ") cursor++; };
  const item = (): void => {
    if (text[cursor++] !== "[") invalid("Expected a table column or selector.");
    let value = "", escaped = false, unescapedSpecial = false;
    while (cursor < text.length && text[cursor] !== "]") {
      const character = text[cursor++]!;
      if (character === "[") invalid("Unexpected nested table column.");
      if (character === "'") { value += text[cursor++]!; escaped = true; }
      else { if (character === "#" || character === "@") unescapedSpecial = true; value += character; }
    }
    if (text[cursor++] !== "]" || !value || value.trim() !== value) invalid("Invalid table column name.");
    const keyword = !escaped ? keywords[value.toLowerCase()] : undefined;
    if (keyword) { if (columns.length) invalid("Table selectors must precede columns."); selectors.push(keyword); }
    else if (unescapedSpecial) unsupported("Unknown table selector or unescaped special column character.");
    else columns.push(value);
  };
  // The @ shorthand is accepted only inside brackets, never as a dynamic-array
  // implicit-intersection operator before an ordinary formula operand.
  if (text.startsWith("[@")) {
    selectors.push("thisRow"); cursor = 2;
    if (text[cursor] === "[") { item(); if (text[cursor] === ":") { cursor++; item(); columnRange = true; } }
    else {
      const nested = "[" + text.slice(cursor);
      const parsed = parseStructuredReference(nested, table);
      if (parsed.selectors.length) invalid("Invalid current-row table column.");
      columns.push(...parsed.columns); cursor = text.length - 1;
    }
    if (text[cursor++] !== "]") invalid("Unclosed current-row table reference.");
  } else if (text === "[]") cursor = text.length;
  else {
    cursor = 1; skip(); const composite = text[cursor] === "["; cursor = composite ? cursor : 0;
    item();
    if (composite) {
      skip();
      while (text[cursor] === "," || text[cursor] === ":") {
        const separator = text[cursor++]!; skip();
        if (separator === ":") { if (columns.length !== 1 || columnRange) invalid("Invalid table column range."); columnRange = true; }
        else if (columns.length) invalid("Discontiguous table columns require separate references.");
        item(); skip();
      }
      if (text[cursor++] !== "]") invalid("Unclosed table reference.");
    }
  }
  if (cursor !== text.length || columns.length > 2 || (columns.length === 2) !== columnRange) invalid("Invalid table specifier.");
  if (new Set(selectors).size !== selectors.length || selectors.length > 2
    || selectors.length === 2 && !(selectors[0] === "headers" && selectors[1] === "data" || selectors[0] === "data" && selectors[1] === "totals")) invalid("Invalid table selector combination.");
  return { kind: "structured", ...(table === undefined ? {} : { table }), selectors, columns };
}

export function printStructuredReference(reference: FormulaStructuredReference): string {
  const columns = reference.columns.map(name => "[" + name.replace(/['\[\]#@]/g, value => "'" + value) + "]").join(":");
  const selectors = reference.selectors.map(selector => "[" + labels[selector] + "]");
  const parts = [...selectors, ...(columns ? [columns] : [])];
  return (reference.table ?? "") + (parts.length === 0 ? "[]" : parts.length === 1 && reference.columns.length < 2 ? parts[0]! : "[" + parts.join(",") + "]");
}
