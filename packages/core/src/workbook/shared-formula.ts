import { checkLimit, invalid } from "../document/validation.js";
import { parseFormula, type FormulaExpression, type FormulaOrigin } from "./formula.js";
import { checkFormulaOrigin, printFormula } from "./formula-print.js";
import type { FormulaRectangle } from "./formula-rewrite.js";

export interface SharedFormulaTable {
  readonly size: number;
  define(id: number, origin: FormulaOrigin, range: FormulaRectangle, source: string): void;
  resolve(id: number, origin: FormulaOrigin): string;
}
/** One bounded derived table per source worksheet, with no per-follower AST. */
export function createSharedFormulaTable(maxGroups = 25_000, maxCodeUnits = 4 * 1024 * 1024): SharedFormulaTable {
  if (![maxGroups, maxCodeUnits].every(value => Number.isSafeInteger(value) && value > 0)) invalid("Invalid shared formula budgets.");
  const groups = new Map<number, { origin: FormulaOrigin; range: FormulaRectangle; source: string; ast: FormulaExpression }>();
  let units = 0;
  const identifier = (id: number): void => { if (!Number.isInteger(id) || id < 0 || id > 0xffff_ffff) invalid("Invalid shared formula group identity."); };
  return Object.freeze({
    get size() { return groups.size; },
    define(id: number, origin: FormulaOrigin, range: FormulaRectangle, source: string): void {
      identifier(id); checkFormulaOrigin(origin); checkFormulaOrigin(range.from); checkFormulaOrigin(range.to);
      if (groups.has(id)) invalid("Repeated shared formula group origin.");
      if (origin.row < range.from.row || origin.row > range.to.row || origin.column < range.from.column || origin.column > range.to.column) invalid("Shared formula origin lies outside its range.");
      const ast = parseFormula(source, origin), nextUnits = units + source.length + JSON.stringify(ast).length;
      checkLimit(groups.size + 1, maxGroups); checkLimit(nextUnits, maxCodeUnits);
      groups.set(id, { origin: { ...origin }, range: { from: { ...range.from }, to: { ...range.to } }, source, ast }); units = nextUnits;
    },
    resolve(id: number, origin: FormulaOrigin): string {
      identifier(id); checkFormulaOrigin(origin);
      const group = groups.get(id); if (!group) invalid("Shared formula follower has no origin.");
      if (origin.row < group.range.from.row || origin.row > group.range.to.row || origin.column < group.range.from.column || origin.column > group.range.to.column) invalid("Shared formula follower lies outside its declared range.");
      if (origin.row === group.origin.row && origin.column === group.origin.column) return group.source;
      const output = printFormula(group.ast, origin); parseFormula(output, origin); return output;
    },
  });
}
