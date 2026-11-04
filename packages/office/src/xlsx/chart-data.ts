import { DocumentError, parseWorkbookFormula, WorkbookErrorValueSchema } from "@dagent-office/core";
import type { WorkbookChartCache, WorkbookChartLabel, WorkbookChartReference, WorkbookFormulaExpression } from "@dagent-office/core";
import { C, check, elements, fail, text, unsupported } from "../ooxml/xml.js";
import type { XmlNode } from "../ooxml/xml.js";
import { attributes, finite, integer, one, spreadsheetText, value } from "./values.js";
import { spreadsheetStringText } from "./strings.js";
import { utf8Bytes } from "./chunks.js";

export interface ChartDataBudget { items: number; bytes: number; readonly maxItems: number; readonly maxBytes: number }
function account(budget: ChartDataBudget, items: number, bytes = 0): void {
  budget.items += items; budget.bytes += bytes;
  if (budget.items > budget.maxItems || budget.bytes > budget.maxBytes) throw new DocumentError("ResourceLimitExceeded", "Chart references and imported caches exceed their cumulative source budget.");
}
function sourceText(node: XmlNode, maximum: number): string {
  const result = spreadsheetStringText(node); if (result.length > maximum) throw new DocumentError("ResourceLimitExceeded", "A chart source string exceeds its text budget."); return result;
}
function cache(source: XmlNode, numeric: boolean, budget: ChartDataBudget): WorkbookChartCache {
  attributes(source, []); check(source, numeric ? ["formatCode", "ptCount", "pt"] : ["ptCount", "pt"], C);
  const count = one(source, "ptCount", C), format = one(source, "formatCode", C), points = elements(source, "pt", C);
  if (count) { attributes(count, ["val"]); check(count, [], C); }
  let size = count ? integer(count, "val") : 0;
  if (!count) for (const point of points) size = Math.max(size, integer(point, "idx") + 1);
  account(budget, size, 64);
  const result: WorkbookChartCache = { state: "imported", count: size, points: [] }, indexes = new Set<number>();
  if (format) { result.numberFormat = sourceText(format, 2048); account(budget, 0, utf8Bytes(JSON.stringify(result.numberFormat))); }
  for (const point of points) {
    attributes(point, numeric ? ["idx", "formatCode"] : ["idx"]); check(point, ["v"], C);
    const index = integer(point, "idx"), v = one(point, "v", C);
    if (index >= size || indexes.has(index) || !v) fail(point, "A chart cache point is duplicated, out of range or missing its value."); indexes.add(index);
    const raw = sourceText(v, 32767);
    const item: WorkbookChartCache["points"][number] = { index, value: numeric ? raw === "" ? null : WorkbookErrorValueSchema.safeParse(raw).success ? raw : finite(raw, v) : raw };
    if (value(point, "formatCode") !== undefined) item.numberFormat = spreadsheetText(value(point, "formatCode")!, point, 2048);
    account(budget, 0, utf8Bytes(JSON.stringify(item)) + 1); result.points.push(item);
  }
  result.points.sort((a, b) => a.index - b.index); return result;
}
export function readChartReference(source: XmlNode, budget: ChartDataBudget): WorkbookChartReference {
  if (source.uri !== C || !["strRef", "numRef"].includes(source.name)) unsupported(source, "Chart vectors must reference worksheet cells or defined names.");
  const numeric = source.name === "numRef", cacheName = numeric ? "numCache" : "strCache";
  attributes(source, []); check(source, ["f", cacheName], C);
  const formula = one(source, "f", C), cached = one(source, cacheName, C); if (!formula) fail(source, "A chart data reference has no formula.");
  attributes(formula, []); if (formula.content.some(item => typeof item !== "string")) fail(formula, "A chart formula must contain only text.");
  // c:f is xsd:string, not ST_Xstring. Do not decode _xNNNN_ sequences in names.
  const raw = text(formula), expression = "=" + raw;
  if (!raw || raw.startsWith("=")) fail(formula, "A native chart reference has invalid formula syntax.");
  const parsed = parseWorkbookFormula(expression, { row: 0, column: 0 });
  const reference = (node: WorkbookFormulaExpression): boolean => ["cell", "range", "axisRange", "name", "structured"].includes(node.kind)
    || node.kind === "error" && node.value === "#REF!" || node.kind === "binary" && [",", " ", ":"].includes(node.operator) && reference(node.left) && reference(node.right)
    || node.kind === "call" && ["INDEX", "OFFSET", "INDIRECT"].includes(node.name);
  if (!reference(parsed)) unsupported(formula, "A chart formula must produce a worksheet reference.");
  account(budget, 1, utf8Bytes(JSON.stringify(expression)) + 64);
  return { expression, valueType: numeric ? "number" : "text", ...(cached ? { cache: cache(cached, numeric, budget) } : {}) };
}
export function readChartVector(source: XmlNode, budget: ChartDataBudget, numericOnly = false): WorkbookChartReference {
  attributes(source, []); check(source, numericOnly ? ["numRef"] : ["strRef", "numRef"], C);
  const references = elements(source, undefined, C); if (references.length !== 1) fail(source, "A chart vector requires exactly one worksheet data reference.");
  return readChartReference(references[0]!, budget);
}
export function readChartLabel(source: XmlNode, budget: ChartDataBudget): WorkbookChartLabel {
  attributes(source, []); check(source, ["strRef", "v"], C);
  const references = elements(source, undefined, C); if (references.length !== 1) fail(source, "A chart label requires exactly one source.");
  const item = references[0]!;
  if (item.name === "strRef") return { kind: "reference", reference: readChartReference(item, budget) };
  const value = sourceText(item, 32767); account(budget, 1, utf8Bytes(JSON.stringify(value)) + 32); return { kind: "text", text: value };
}
