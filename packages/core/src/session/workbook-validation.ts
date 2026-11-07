import type { DeepReadonly, TransactionPrepareOptions, WorkbookModel, WorkbookRangeChange } from "../runtime.js";
import type { WorkbookDataValidation } from "../contracts/workbook.js";
import type { WorkbookReferenceLocation } from "../workbook/reference.js";
import type { WorkbookRuleValue, WorkbookValidationEvaluation, WorkbookRuleIssue } from "../workbook/rule-reader.js";
import { DEFAULT_WORKBOOK_RULE_LIMITS } from "../workbook/rule-reader.js";
import { createWorkbookScalarKernel } from "../workbook/scalar.js";
import { createWorkbookCalculator } from "../workbook/calculation.js";
import { beginWorkbookValidation, workbookRuleIssue } from "../workbook/validation.js";
import { scalarToText } from "../workbook/scalar-values.js";
import { checkLimit } from "../document/validation.js";
import { DocumentError } from "../errors.js";
import { rangeContains } from "./workbook-cells.js";
import { workbookExpressionDependencies } from "../workbook/expression-dependencies.js";
import { parseFormula } from "../workbook/formula.js";
import { WorkbookDependencyIndex } from "../workbook/dependency-index.js";

/** All cell-writing entry points validate the final candidate through this
 * shared rule implementation. Unsupported stop conditions never mean pass. */
export function validateWorkbookCandidate(previous: WorkbookModel, model: WorkbookModel, sourceRanges: readonly WorkbookRangeChange[], options?: TransactionPrepareOptions): void {
  const snapshot = model.snapshot(), rules = new Map(snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, sheet.dataValidations ?? []]));
  const prior = previous.snapshot(), previousRules = new Map(prior.metadata.workbook.sheets.flatMap(sheet => sheet.dataValidations?.map(rule => [rule.ruleId, rule] as const) ?? []));
  const contextChanged = JSON.stringify(prior.metadata.workbook.names) !== JSON.stringify(snapshot.metadata.workbook.names) || prior.metadata.locale !== snapshot.metadata.locale;
  const changedRules = new Set([...rules.values()].flatMap(items => items.filter(rule => contextChanged || JSON.stringify(previousRules.get(rule.ruleId)) !== JSON.stringify(rule)).map(rule => rule.ruleId)));
  if (![...rules.values()].some(items => items.length)) return;
  const previousSheets = new Set(prior.metadata.workbook.sheets.map(sheet => sheet.sheetId)), nextSheets = new Set(rules.keys());
  const changes: WorkbookRangeChange[] = [];
  for (const change of sourceRanges) for (let row = change.range.from.row; row <= change.range.to.row; row++) for (let column = change.range.from.column; column <= change.range.to.column; column++) {
    const before = previousSheets.has(change.sheetId) ? previous.getCell(change.sheetId, row, column) : undefined, after = nextSheets.has(change.sheetId) ? model.getCell(change.sheetId, row, column) : undefined;
    const signature = (cell: typeof before): string => cell?.formula ? cell.formula.expression : JSON.stringify(cell?.value ?? { type: "blank" });
    if (signature(before) !== signature(after)) changes.push({ sheetId: change.sheetId, range: { from: { row, column }, to: { row, column } } });
  }
  if (![...rules.values()].some(items => items.length) || !changes.length && !changedRules.size) return;
  const limits = DEFAULT_WORKBOOK_RULE_LIMITS;
  let steps = 0, units = 0, cells = 0;
  const step = (): void => checkLimit(++steps, limits.maxSteps);
  const account = (amount: number): void => { units += amount; checkLimit(units, limits.maxCodeUnits); };
  let currentModel = previous;
  const calculator = createWorkbookCalculator(() => currentModel, { maxSteps: limits.maxSteps, maxReadCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits });
  const beforeCalculation = calculator.calculateSync({ ...(options?.calculationContext ? { context: options.calculationContext } : {}) });
  steps += beforeCalculation.steps; account(beforeCalculation.codeUnits); let readCells = beforeCalculation.readCells;
  currentModel = model;
  const calculation = calculator.calculateSync({ ...(options?.calculationContext ? { context: options.calculationContext } : {}) });
  steps += calculation.steps; checkLimit(steps, limits.maxSteps); account(calculation.codeUnits); readCells += calculation.readCells; checkLimit(readCells, limits.maxValues);
  const cellKey = (at: WorkbookReferenceLocation): string => JSON.stringify([at.sheetId, at.row, at.column]);
  const beforeValues = new Map(beforeCalculation.cells.map(cell => [cellKey(cell), cell.state === "valid" ? JSON.stringify(cell.value) : cell.state === "unavailable" ? JSON.stringify(cell.issue) : "stale"]));
  const derived = calculation.cells.filter(cell => beforeValues.get(cellKey(cell)) !== (cell.state === "valid" ? JSON.stringify(cell.value) : cell.state === "unavailable" ? JSON.stringify(cell.issue) : "stale"));
  const affected = new WorkbookDependencyIndex([...changes.map(change => ({ area: { sheetId: change.sheetId, ...change.range }, owner: "source" })), ...derived.map(cell => ({ area: { sheetId: cell.sheetId, from: cell, to: cell }, owner: "calculated" }))], step);
  const evaluator = createWorkbookScalarKernel(model, { maxSteps: limits.maxSteps, maxCells: limits.maxValues, maxCodeUnits: limits.maxCodeUnits }, options?.calculationContext, {
    step, account, read: () => checkLimit(++readCells, limits.maxValues),
    cachedCell: location => {
      const value = calculator.getCell(location); return value.state === "valid" ? value.value : value.state === "unavailable" ? value.issue : undefined;
    },
  });
  function current(location: WorkbookReferenceLocation): WorkbookRuleValue {
    step(); const value = evaluator.cell(location); return value.kind === "value" ? { kind: "value", value: value.value } : { kind: "unavailable", reason: value.reason, ...(value.functionName ? { functionName: value.functionName } : {}) };
  }
  function validate(rule: DeepReadonly<WorkbookDataValidation>, location: WorkbookReferenceLocation): WorkbookValidationEvaluation {
    const plan = beginWorkbookValidation(rule, location, current(location), { account,
      resolve: (expression, at) => model.resolveReference(expression, { sheetId: location.sheetId, ...at }),
      evaluate: (expression, definition, coordinate) => {
        const result = evaluator.evaluate(expression, { sheetId: location.sheetId, ...definition.origin }, { at: { sheetId: location.sheetId, ...coordinate } });
        return result.kind === "value" ? { kind: "value", value: result.value } : { kind: "unavailable", reason: result.reason, ...(result.functionName ? { functionName: result.functionName } : {}) };
      },
    });
    if (!("kind" in plan)) return plan;
    let unknown: WorkbookRuleIssue | undefined, count = 0;
    for (const area of plan.areas) for (const chunk of snapshot.manifest.chunks) {
      step(); if (chunk.sheetId !== area.sheetId) continue;
      const row = Math.max(area.from.row, chunk.row), lastRow = Math.min(area.to.row, chunk.row + 255), column = Math.max(area.from.column, chunk.column), lastColumn = Math.min(area.to.column, chunk.column + 31);
      if (row > lastRow || column > lastColumn) continue;
      const columns = lastColumn - column + 1, batchRows = Math.max(1, Math.floor(4096 / columns));
      for (let start = row; start <= lastRow; start += batchRows) for (const cell of model.readWindow({ sheetId: area.sheetId, row: start, column, rows: Math.min(batchRows, lastRow - start + 1), columns })) {
        step(); checkLimit(++count, limits.maxValues); const value = current({ sheetId: area.sheetId, row: cell.row, column: cell.column });
        if (value.kind === "unavailable") { unknown = workbookRuleIssue(value); continue; }
        const text = scalarToText(value.value); if (typeof text === "string" && text.toUpperCase() === plan.text.toUpperCase()) return { ruleId: plan.ruleId, state: "valid" };
      }
    }
    const area = plan.areas[0]!, total = (area.to.row - area.from.row + 1) * (area.to.column - area.from.column + 1);
    if (plan.text === "" && count < total) return { ruleId: plan.ruleId, state: "valid" };
    return unknown ? { ruleId: plan.ruleId, state: "unavailable", issue: unknown } : { ruleId: plan.ruleId, state: "invalid" };
  }
  try {
    const targets = new Map<string, { location: WorkbookReferenceLocation; rule: DeepReadonly<WorkbookDataValidation> }>();
    function relevant(rule: DeepReadonly<WorkbookDataValidation>, location: WorkbookReferenceLocation): boolean {
      if (rule.kind === "none") return false;
      if (changedRules.has(rule.ruleId) || affected.query({ sheetId: location.sheetId, from: location, to: location }, step).next().done === false) return true;
      let match = false;
      const expressions = "expression" in rule ? [rule.expression] : rule.formulas;
      for (const expression of expressions) workbookExpressionDependencies(model, parseFormula(expression, rule.origin), location, step, area => { if (affected.query(area, step).next().done === false) match = true; });
      return match;
    }
    function add(location: WorkbookReferenceLocation): void {
      const rule = rules.get(location.sheetId)?.find(rule => rule.ranges.some(range => rangeContains(range, location.row, location.column))); if (!rule) return;
      const key = cellKey(location); if (targets.has(key)) return;
      if (!relevant(rule, location)) return;
      checkLimit(targets.size + 1, limits.maxCells); targets.set(key, { location, rule });
    }
    // Explicit clears must include now-absent cells. Other implicit blanks are
    // skipped only when the definition itself explicitly permits blank values.
    for (const change of changes) for (let row = change.range.from.row; row <= change.range.to.row; row++) for (let column = change.range.from.column; column <= change.range.to.column; column++) { step(); add({ sheetId: change.sheetId, row, column }); }
    for (const chunk of snapshot.manifest.chunks) {
      if (!rules.get(chunk.sheetId)?.length) continue;
      for (let row = chunk.row; row < chunk.row + 256; row += 128) for (const cell of model.readWindow({ sheetId: chunk.sheetId, row, column: chunk.column, rows: 128, columns: 32 })) { step(); add({ sheetId: chunk.sheetId, row: cell.row, column: cell.column }); }
    }
    for (const [sheetId, definitions] of rules) for (const rule of definitions) {
      checkLimit(definitions.length, limits.maxRules);
      if (rule.allowBlank || rule.kind === "none") continue;
      const expressions = "expression" in rule ? [rule.expression] : rule.formulas;
      let dependent = false;
      for (const expression of expressions) workbookExpressionDependencies(model, parseFormula(expression, rule.origin), { sheetId, ...rule.origin }, step, () => { dependent = true; });
      if (!changedRules.has(rule.ruleId) && !dependent) continue;
      for (const range of rule.ranges) {
        checkLimit((range.to.row - range.from.row + 1) * (range.to.column - range.from.column + 1), limits.maxCells);
        for (let row = range.from.row; row <= range.to.row; row++) for (let column = range.from.column; column <= range.to.column; column++) { step(); add({ sheetId, row, column }); }
      }
    }
    for (const { location, rule } of targets.values()) {
      step(); const { sheetId, row, column } = location; checkLimit(++cells, limits.maxCells);
      const result = validate(rule, location);
      if (result.state === "valid" || result.state === "none") continue;
      if (result.state === "invalid" && rule.errorStyle !== "stop" && options?.validation?.acceptedRuleIds.includes(rule.ruleId)) continue;
      throw new DocumentError("InvalidOperation", result.state === "unavailable" ? "The candidate's data validation cannot be evaluated." : "The candidate does not satisfy its data validation.", [
        { code: rule.errorStyle === "stop" || result.state === "unavailable" ? "validation-blocked" : "validation-confirmation-required", path: [sheetId, row, column, rule.ruleId], message: rule.error || result.issue?.reason || "The cell does not satisfy the rule." },
      ]);
    }
  } finally { evaluator.dispose(); calculator.dispose(); }
}
