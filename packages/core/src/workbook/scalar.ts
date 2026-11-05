import { DocumentError } from "../errors.js";
import type { DeepReadonly, WorkbookCell, WorkbookModel } from "../runtime.js";
import type { WorkbookDefinedName } from "../contracts/workbook.js";
import { checkLimit, freeze } from "../document/validation.js";
import { parseFormula } from "./formula.js";
import type { FormulaExpression, FormulaOrigin } from "./formula.js";
import { printFormula } from "./formula-print.js";
import type { WorkbookReferenceArea, WorkbookReferenceLocation } from "./reference.js";
import { scalarBinary, scalarBlank, scalarError, scalarNumber, scalarToNumber } from "./scalar-values.js";
import type { ScalarValue } from "./scalar-values.js";
import { evaluateScalarFunction } from "./scalar-functions.js";

export type WorkbookScalarUnavailable = { readonly kind: "unavailable"; readonly reason: "unsupported-function" | "circular-reference" | "circular-name" | "ambiguous-name-scope" | "calculation-required"; readonly functionName?: string };
export type WorkbookScalarResult = ({ readonly kind: "value"; readonly value: ScalarValue } | WorkbookScalarUnavailable) & {
  readonly documentId: string; readonly revision: number; readonly steps: number; readonly readCells: number;
};
export interface WorkbookScalarLimits {
  readonly maxSteps: number; readonly maxCells: number; readonly maxDepth: number; readonly maxCodeUnits: number; readonly maxWindowCells: number;
}
export interface WorkbookScalarQueryOptions {
  /** Shift a parsed template without editing its source expression. */
  readonly at?: WorkbookReferenceLocation;
  readonly signal?: { readonly aborted: boolean };
}
export interface WorkbookScalarEvaluator {
  evaluate(expression: string, origin: WorkbookReferenceLocation, options?: WorkbookScalarQueryOptions): WorkbookScalarResult;
  cell(location: WorkbookReferenceLocation, options?: Pick<WorkbookScalarQueryOptions, "signal">): WorkbookScalarResult;
  dispose(): void;
}
export type ScalarEvaluation = ScalarValue | WorkbookScalarUnavailable | { readonly kind: "references"; readonly areas: readonly DeepReadonly<WorkbookReferenceArea>[] };
export type ScalarResolved = ScalarValue | WorkbookScalarUnavailable;
export interface ScalarEntry { readonly value: ScalarResolved; readonly count: number; readonly reference: boolean; readonly hasContent: boolean }
export interface ScalarFunctionContext {
  readonly origin: WorkbookReferenceLocation;
  visit(node: FormulaExpression | null): ScalarEvaluation;
  scalar(value: ScalarEvaluation): ScalarResolved;
  entries(value: ScalarEvaluation, presenceOnly?: boolean): Generator<ScalarEntry>;
  step(): void;
}
const defaults: WorkbookScalarLimits = Object.freeze({ maxSteps: 1_000_000, maxCells: 250_000, maxDepth: 128, maxCodeUnits: 8 * 1024 * 1024, maxWindowCells: 4096 });
type SourceCell = Pick<WorkbookCell, "value" | "formula">;
type Name = DeepReadonly<WorkbookDefinedName>;
const unavailable = (reason: WorkbookScalarUnavailable["reason"]): WorkbookScalarUnavailable => ({ kind: "unavailable", reason });

/** A bounded, read-only expression kernel shared by rule reading and later
 * calculation. It never accepts imported caches as freshly evaluated values,
 * changes source/history, reads a clock or expands an entire sparse axis. */
export function createWorkbookScalarEvaluator(model: WorkbookModel, input: Partial<WorkbookScalarLimits> = {}): WorkbookScalarEvaluator {
  if (Object.keys(input).some(key => !Object.hasOwn(defaults, key)) || !Object.values(input).every(value => Number.isSafeInteger(value) && value! > 0))
    throw new DocumentError("InvalidOperation", "Scalar evaluation requires supported positive safe-integer budgets.");
  const limits = { ...defaults, ...input }, snapshot = model.snapshot(), workbook = snapshot.metadata.workbook;
  if (limits.maxDepth > 128) throw new DocumentError("InvalidOperation", "This scalar kernel supports a maximum evaluation depth of 128.");
  let sourceModel: WorkbookModel | undefined = model;
  const documentId = snapshot.manifest.documentId, revision = snapshot.manifest.revision;
  const sheets = new Map(workbook.sheets.map((sheet, index) => [sheet.sheetId, { sheet, index }])), sheetNames = new Map(workbook.sheets.map(sheet => [sheet.name.toUpperCase(), sheet.sheetId]));
  const names = new Map((workbook.names ?? []).map(name => [JSON.stringify([name.sheetId ?? null, name.name.toUpperCase()]), name]));
  const localNames = new Set((workbook.names ?? []).filter(name => name.sheetId !== undefined).map(name => name.name.toUpperCase()));
  const chunks = new Map<string, (typeof snapshot.manifest.chunks)[number][]>(workbook.sheets.map(sheet => [sheet.sheetId, []]));
  for (const chunk of snapshot.manifest.chunks) chunks.get(chunk.sheetId)!.push(chunk);
  const parsed = new Map<string, { readonly ast: FormulaExpression; readonly units: number }>(); let parseUnits = 0, disposed = false;
  const checkLocation = (value: WorkbookReferenceLocation): void => {
    if (disposed) throw new DocumentError("InvalidOperation", "This scalar evaluator has been disposed.");
    if (!value || typeof value !== "object" || Object.keys(value).some(key => !["sheetId", "row", "column"].includes(key)) || !sheets.has(value.sheetId)
      || !Number.isSafeInteger(value.row) || !Number.isSafeInteger(value.column) || value.row < 0 || value.row >= 1_048_576 || value.column < 0 || value.column >= 16_384)
      throw new DocumentError("InvalidOperation", "Scalar evaluation requires an existing worksheet coordinate.");
  };
  const parse = (expression: string, origin: FormulaOrigin): FormulaExpression => {
    const key = JSON.stringify([expression, origin.row, origin.column]), cached = parsed.get(key); if (cached) return cached.ast;
    const ast = parseFormula(expression, origin), units = key.length + expression.length;
    while (parsed.size && (parsed.size >= 256 || parseUnits + units > 2 * 1024 * 1024)) { const first = parsed.keys().next().value!; parseUnits -= parsed.get(first)!.units; parsed.delete(first); }
    if (units <= 2 * 1024 * 1024) { parsed.set(key, { ast, units }); parseUnits += units; }
    return ast;
  };
  function query(start: (visit: (node: FormulaExpression, location: WorkbookReferenceLocation, depth: number, nameContext?: Name) => ScalarEvaluation,
    readCell: (location: WorkbookReferenceLocation, depth: number) => ScalarResolved) => ScalarEvaluation, location: WorkbookReferenceLocation, options: WorkbookScalarQueryOptions, initialUnits = 0): WorkbookScalarResult {
    checkLocation(location);
    if (Object.keys(options).some(key => !["at", "signal"].includes(key))) throw new DocumentError("InvalidOperation", "Scalar query received an unsupported option.");
    const memo = new Map<number, ScalarResolved>(), sources = new Map<number, SourceCell | undefined>(), path = new Set<number>(), namePath = new Map<string, Set<number>>();
    let steps = 0, units = initialUnits, readCells = 0; checkLimit(units, limits.maxCodeUnits);
    const step = (depth = 0): void => {
      if (options.signal?.aborted) throw new DocumentError("OperationCancelled", "Scalar evaluation was cancelled.");
      checkLimit(++steps, limits.maxSteps); checkLimit(depth, limits.maxDepth);
    };
    const account = (amount: number): void => { units += amount; checkLimit(units, limits.maxCodeUnits); };
    const key = (at: WorkbookReferenceLocation): number => sheets.get(at.sheetId)!.index * 17_179_869_184 + at.row * 16_384 + at.column;
    function remember(at: WorkbookReferenceLocation, cell: SourceCell | undefined): void {
      const id = key(at); if (sources.has(id)) return;
      checkLimit(++readCells, limits.maxCells);
      account((cell?.formula?.expression.length ?? 0) + (cell?.value.type === "text" ? cell.value.value.length : 0));
      sources.set(id, cell ? { value: cell.value, ...(cell.formula ? { formula: cell.formula } : {}) } : undefined);
    }
    function readCell(at: WorkbookReferenceLocation, depth: number): ScalarResolved {
      step(depth); const id = key(at), cached = memo.get(id); if (cached) return cached;
      if (path.has(id)) return unavailable("circular-reference");
      if (!sources.has(id)) remember(at, sourceModel!.getCell(at.sheetId, at.row, at.column));
      const cell = sources.get(id);
      if (!cell?.formula) { const value = cell?.value ?? scalarBlank; memo.set(id, value); return value; }
      path.add(id);
      try {
        const result = scalar(visit(parse(cell.formula.expression, at), at, depth + 1), at, depth + 1);
        const value = "type" in result && result.type === "blank" ? scalarNumber(0) : result;
        memo.set(id, value); return value;
      } finally { path.delete(id); }
    }
    function scalar(value: ScalarEvaluation, at: WorkbookReferenceLocation, depth: number): ScalarResolved {
      if ("type" in value || value.kind === "unavailable") return value;
      if (value.areas.length !== 1) return scalarError("#VALUE!");
      const area = value.areas[0]!, row = area.from.row === area.to.row ? area.from.row : at.row, column = area.from.column === area.to.column ? area.from.column : at.column;
      if (area.from.row !== area.to.row && area.from.column !== area.to.column || row < area.from.row || row > area.to.row || column < area.from.column || column > area.to.column) return scalarError("#VALUE!");
      return readCell({ sheetId: area.sheetId, row, column }, depth + 1);
    }
    function* entries(value: ScalarEvaluation, depth: number, presenceOnly = false): Generator<ScalarEntry> {
      if ("type" in value || value.kind === "unavailable") { yield { value, count: 1, reference: false, hasContent: !("type" in value) || value.type !== "blank" }; return; }
      for (const area of value.areas) {
        let stored = 0;
        for (const chunk of chunks.get(area.sheetId) ?? []) {
          step(depth);
          const firstRow = Math.max(area.from.row, chunk.row), lastRow = Math.min(area.to.row, chunk.row + 255), firstColumn = Math.max(area.from.column, chunk.column), lastColumn = Math.min(area.to.column, chunk.column + 31);
          if (firstRow > lastRow || firstColumn > lastColumn) continue;
          for (let column = firstColumn; column <= lastColumn;) {
            const columns = Math.min(lastColumn - column + 1, limits.maxWindowCells), rows = Math.max(1, Math.floor(limits.maxWindowCells / columns));
            for (let row = firstRow; row <= lastRow; row += rows) {
              step(depth);
              const cells = sourceModel!.readWindow({ sheetId: area.sheetId, row, column, rows: Math.min(rows, lastRow - row + 1), columns });
              for (const cell of cells) {
                step(depth); stored++; const location = { sheetId: area.sheetId, row: cell.row, column: cell.column }; remember(location, cell);
                yield { value: presenceOnly ? scalarBlank : readCell(location, depth + 1), count: 1, reference: true, hasContent: !!cell.formula || cell.value.type !== "blank" };
              }
            }
            column += columns;
          }
        }
        const missing = (area.to.row - area.from.row + 1) * (area.to.column - area.from.column + 1) - stored;
        if (missing) yield { value: scalarBlank, count: missing, reference: true, hasContent: false };
      }
    }
    function reference(node: FormulaExpression, at: WorkbookReferenceLocation): ScalarEvaluation {
      const result = sourceModel!.resolveReference(printFormula(node, at), at);
      return result.kind === "ranges" ? { kind: "references", areas: result.areas } : result.kind === "error" ? scalarError(result.code) : unavailable(result.reason);
    }
    function visit(node: FormulaExpression, at: WorkbookReferenceLocation, depth: number, nameContext?: Name): ScalarEvaluation {
      step(depth);
      if (node.kind === "number") return scalarNumber(node.value);
      if (node.kind === "string") { account(node.value.length); return { type: "text", value: node.value }; }
      if (node.kind === "boolean") return { type: "boolean", value: node.value };
      if (node.kind === "error") return scalarError(node.value as Extract<ScalarValue, { type: "error" }>["value"]);
      if (node.kind === "name") {
        if (node.sheet === undefined && nameContext?.sheetId === undefined && nameContext && localNames.has(node.name.toUpperCase())) return unavailable("ambiguous-name-scope");
        const explicit = node.sheet === undefined ? undefined : sheetNames.get(node.sheet.toUpperCase()); if (node.sheet !== undefined && !explicit) return scalarError("#REF!");
        const scope = explicit ?? nameContext?.sheetId ?? at.sheetId;
        const definition = names.get(JSON.stringify([scope, node.name.toUpperCase()])) ?? (explicit ? undefined : names.get(JSON.stringify([null, node.name.toUpperCase()])));
        if (!definition) return reference(node, at);
        const id = key(at), active = namePath.get(definition.nameId) ?? new Set<number>();
        if (active.has(id)) return unavailable("circular-name");
        active.add(id); namePath.set(definition.nameId, active);
        try { return visit(parse(definition.expression, definition.origin), { ...at, sheetId: definition.sheetId ?? at.sheetId }, depth + 1, definition); }
        finally { active.delete(id); }
      }
      if (node.kind === "cell" || node.kind === "range" || node.kind === "axisRange" || node.kind === "structured"
        || node.kind === "binary" && [",", " ", ":"].includes(node.operator)) return reference(node, at);
      if (node.kind === "unary") {
        const value = scalar(visit(node.operand, at, depth + 1, nameContext), at, depth + 1); if (!("type" in value)) return value;
        const number = scalarToNumber(value); return typeof number === "number" ? scalarNumber(node.operator === "-" ? -number : node.operator === "%" ? number / 100 : number) : number;
      }
      if (node.kind === "binary") {
        const a = scalar(visit(node.left, at, depth + 1, nameContext), at, depth + 1); if (!("type" in a)) return a;
        const b = scalar(visit(node.right, at, depth + 1, nameContext), at, depth + 1); if (!("type" in b)) return b;
        const result = scalarBinary(node.operator, a, b); if (result.type === "text") account(result.value.length); return result;
      }
      return evaluateScalarFunction(node, { origin: at, step: () => step(depth + 1), visit: argument => argument ? visit(argument, at, depth + 1, nameContext) : scalarNumber(0),
        scalar: value => scalar(value, at, depth + 1), entries: (value, presenceOnly) => entries(value, depth + 1, presenceOnly) });
    }
    step(); const result = scalar(start(visit, readCell), location, 0), identity = { documentId, revision, steps, readCells };
    return freeze("type" in result ? { ...identity, kind: "value", value: result.type === "blank" ? scalarNumber(0) : result } : { ...identity, ...result });
  }
  return Object.freeze({
    evaluate(expression: string, origin: WorkbookReferenceLocation, options: WorkbookScalarQueryOptions = {}) {
      if (typeof expression !== "string" || !options || typeof options !== "object") throw new DocumentError("InvalidOperation", "Scalar evaluation requires an expression and query options.");
      checkLocation(origin); const location = options.at ?? origin; checkLocation(location);
      return query(visit => visit(parse(expression, origin), location, 0), location, options, expression.length);
    },
    cell(location: WorkbookReferenceLocation, options: Pick<WorkbookScalarQueryOptions, "signal"> = {}) {
      if (!options || typeof options !== "object" || Object.keys(options).some(key => key !== "signal")) throw new DocumentError("InvalidOperation", "Scalar cell query received unsupported options.");
      return query((_visit, readCell) => readCell(location, 0), location, options);
    },
    dispose() { disposed = true; sourceModel = undefined; parsed.clear(); parseUnits = 0; sheets.clear(); sheetNames.clear(); names.clear(); localNames.clear(); chunks.clear(); },
  });
}
