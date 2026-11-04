import { invalid } from "../document/validation.js";
import { checkFormulaOrigin } from "./formula-print.js";
import type { FormulaLocation, FormulaRectangle } from "./formula-rewrite.js";
import type { FormulaStructuredReference } from "./structured-reference.js";

/** Format-neutral source geometry used by import, dependencies and transactions. */
export interface FormulaTableDefinition {
  readonly name: string;
  readonly sheet: string;
  readonly range: FormulaRectangle;
  readonly header: boolean;
  readonly totals: boolean;
  readonly columns: readonly string[];
}
export type BoundTableReference = { readonly sheet: string; readonly range?: FormulaRectangle } | { readonly error: "#REF!" | "#VALUE!" | "#NAME?" };

/** Resolve at most one rectangle, including an explicitly empty data body. */
export function bindStructuredReference(reference: FormulaStructuredReference, origin: FormulaLocation, tables: readonly FormulaTableDefinition[], containingTable?: string): BoundTableReference {
  const name = reference.table ?? containingTable;
  if (!name) invalid("An unqualified table reference requires its containing table.");
  const matches = tables.filter(table => table.name.toLowerCase() === name.toLowerCase());
  if (matches.length > 1) invalid("Table names must be unique ignoring case.");
  const table = matches[0]; if (!table) return { error: "#NAME?" };
  checkFormulaOrigin(table.range.from); checkFormulaOrigin(table.range.to); checkFormulaOrigin(origin);
  const { from, to } = table.range;
  if (from.row > to.row || from.column > to.column || table.columns.length !== to.column - from.column + 1
    || Number(table.header) + Number(table.totals) > to.row - from.row + 1
    || new Set(table.columns.map(column => column.toLowerCase())).size !== table.columns.length) invalid("Invalid table reference geometry.");
  const indexes = reference.columns.map(column => table.columns.findIndex(candidate => candidate.toLowerCase() === column.toLowerCase()));
  if (indexes.some(index => index < 0)) return { error: "#REF!" };
  const firstColumn = from.column + (indexes.length ? Math.min(...indexes) : 0), lastColumn = from.column + (indexes.length ? Math.max(...indexes) : table.columns.length - 1);
  const firstData = from.row + Number(table.header), lastData = to.row - Number(table.totals);
  const selectors = reference.selectors.length ? reference.selectors : ["data"];
  let first = to.row + 1, last = from.row - 1;
  for (const selector of selectors) {
    let start: number, end: number;
    if (selector === "all") { start = from.row; end = to.row; }
    else if (selector === "headers") { if (!table.header) continue; start = end = from.row; }
    else if (selector === "totals") { if (!table.totals) continue; start = end = to.row; }
    else if (selector === "thisRow") {
      if (origin.row < firstData || origin.row > lastData) return { error: "#VALUE!" };
      start = end = origin.row;
    } else { start = firstData; end = lastData; }
    if (start <= end) { first = Math.min(first, start); last = Math.max(last, end); }
  }
  return { sheet: table.sheet, ...(first <= last ? { range: { from: { row: first, column: firstColumn }, to: { row: last, column: lastColumn } } } : {}) };
}
