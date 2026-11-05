import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookChart, WorkbookChartVector, WorkbookDisplayFormat, WorkbookNumberFormatter, WorkbookSnapshot, WorkbookValue } from "@dagent-office/core";
import type { BasicChartData } from "../chart.js";
import { checkWindowIdentity } from "./window-source.js";
import type { WorkbookChartCell, WorkbookChartData, WorkbookChartReadResponse, WorkbookChartReferenceData, WorkbookChartReferenceSlot } from "./chart-source.js";

export interface WorkbookChartBindingLimits { readonly maxCharts: number; readonly maxReferences: number; readonly maxPoints: number; readonly maxCodeUnits: number }
export interface WorkbookChartDataDiagnostic {
  readonly code: "reference-error" | "unresolved-reference" | "non-scalar-label" | "stale-cell" | "unavailable-cell" | "imported-cell" | "error-point";
  readonly chartId: string;
  readonly role: WorkbookChartReferenceSlot["role"];
  readonly seriesId?: string;
  readonly index?: number;
  readonly detail?: string;
  readonly issue?: WorkbookChartCell["issue"];
}
export type WorkbookChartBinding = {
  readonly chartId: string;
  readonly diagnostics: readonly WorkbookChartDataDiagnostic[];
} & ({
  readonly kind: "available";
  readonly data: BasicChartData;
  readonly titles: Readonly<Partial<Record<"title" | "category-axis-title" | "value-axis-title", string>>>;
  /** Same displayed positions as data.series; formats are current-cell or
   * native-literal definitions, never imported reference-cache formats. */
  readonly valueFormats: readonly (readonly (WorkbookDisplayFormat | undefined)[])[];
  /** Numeric category source values for explicit axis-format overrides. Text
   * categories retain their already formatted current-source strings. */
  readonly categoryNumbers: readonly (number | null)[];
  /** First currently plotted numeric source, in series/source order. */
  readonly valueAxisFormat?: WorkbookDisplayFormat;
} | { readonly kind: "unavailable" });
interface Point {
  readonly value: DeepReadonly<WorkbookValue>;
  readonly hidden: boolean;
  readonly state: "constant" | "imported" | "valid" | "stale" | "unavailable";
  readonly format?: WorkbookDisplayFormat;
  readonly issue?: WorkbookChartCell["issue"];
}
type Role = WorkbookChartReferenceSlot["role"];
const slotKey = (role: Role, seriesId?: string): string => `${role}:${seriesId ?? ""}`;

/** Bind current typed values to plotting positions. Hidden data, absent data
 * and errors stay distinct; removing a source position retains its original
 * index for point styling. Reference caches are never consulted, including
 * when a current reference is unresolved or a formula result is stale. */
export function bindWorkbookChartData(snapshot: WorkbookSnapshot, response: WorkbookChartReadResponse, formatter: WorkbookNumberFormatter,
  input: Partial<WorkbookChartBindingLimits> = {}): readonly WorkbookChartBinding[] {
  checkWindowIdentity(response, snapshot);
  const limits = { maxCharts: 128, maxReferences: 5000, maxPoints: 50_000, maxCodeUnits: 8 * 1024 * 1024, ...input };
  if (Object.keys(input).some(key => !["maxCharts", "maxReferences", "maxPoints", "maxCodeUnits"].includes(key)) || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0))
    throw new DocumentError("InvalidOperation", "Workbook chart binding requires positive supported budgets.");
  if (response.charts.length > limits.maxCharts) throw new DocumentError("ResourceLimitExceeded", "Workbook chart binding exceeds its chart budget.");
  const charts = new Map((snapshot.metadata.workbook.charts ?? []).map(chart => [chart.chartId, chart]));
  const owners = new Map((snapshot.metadata.workbook.objects ?? []).filter(object => object.kind === "chart").map(object => [object.chartId, object.sheetId]));
  const formats = new Map(snapshot.metadata.workbook.numberFormats.map(format => [format.formatId, format]));
  let points = 0, derivedPoints = 0, units = 0, references = 0;
  const accountPoints = (count: number, derived = false): void => {
    if (derived) derivedPoints += count; else points += count;
    if (!Number.isSafeInteger(count) || count < 0 || points > limits.maxPoints || derivedPoints > limits.maxPoints) throw new DocumentError("ResourceLimitExceeded", "Workbook chart binding exceeds its cumulative point budget.");
  };
  const accountText = (text: string): string => { units += text.length; if (units > limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook chart binding exceeds its cumulative text budget."); return text; };
  const seen = new Set<string>();
  for (const data of response.charts) {
    if (seen.has(data.chartId) || !charts.has(data.chartId) || owners.get(data.chartId) !== data.sheetId) throw new DocumentError("InvalidDocument", "Chart data does not uniquely match its owning workbook object.");
    seen.add(data.chartId);
    references += data.references.length;
    if (references > limits.maxReferences) throw new DocumentError("ResourceLimitExceeded", "Workbook chart binding exceeds its reference budget.");
    for (const reference of data.references) {
      if (reference.data.kind === "literal") accountPoints(reference.data.count);
      if (reference.data.kind === "cells") for (const area of reference.data.areas) accountPoints(area.cells.length);
    }
  }
  const vector = (data: WorkbookChartReferenceData): readonly Point[] => {
    if (data.kind === "literal") {
      const result: Point[] = Array.from({ length: data.count }, () => ({ value: { type: "blank" }, hidden: false, state: "constant" }));
      for (const point of data.points) {
        if (!Number.isSafeInteger(point.index) || point.index < 0 || point.index >= data.count) throw new DocumentError("InvalidDocument", "A literal chart point lies outside its vector.");
        const value: WorkbookValue = point.value === null ? { type: "blank" } : typeof point.value === "number" ? { type: "number", value: point.value }
          : data.valueType === "text" ? { type: "text", value: accountText(point.value) } : { type: "error", value: accountText(point.value) as Extract<WorkbookValue, { type: "error" }>["value"] };
        const code = point.numberFormat ?? data.numberFormat;
        result[point.index] = { value, hidden: false, state: "constant", ...(code === undefined ? {} : { format: Object.freeze({ code: accountText(code) }) }) };
      }
      return result;
    }
    if (data.kind !== "cells") return [];
    return data.areas.flatMap(area => area.cells.map(cell => {
      if (cell.value.type === "text" || cell.value.type === "error") accountText(cell.value.value);
      const format = cell.numberFormatId === undefined ? undefined : formats.get(cell.numberFormatId);
      if (cell.numberFormatId !== undefined && !format) throw new DocumentError("InvalidDocument", "A current chart cell references an unknown number format.");
      let issue: WorkbookChartCell["issue"];
      if (cell.issue !== undefined) {
        if (!cell.issue || cell.state !== "unavailable" || Object.keys(cell.issue).some(key => !["reason", "functionName"].includes(key))
          || !["unsupported-function", "circular-reference", "circular-name", "ambiguous-name-scope", "calculation-required"].includes(cell.issue.reason)
          || cell.issue.functionName !== undefined && (typeof cell.issue.functionName !== "string" || !cell.issue.functionName.length))
          throw new DocumentError("InvalidDocument", "A chart cell has an invalid calculation issue.");
        issue = Object.freeze({ reason: accountText(cell.issue.reason) as NonNullable<WorkbookChartCell["issue"]>["reason"],
          ...(cell.issue.functionName !== undefined ? { functionName: accountText(cell.issue.functionName) } : {}) });
      }
      return { value: cell.value, hidden: cell.rowHidden || cell.columnHidden, state: cell.state, ...(format ? { format } : {}), ...(issue ? { issue } : {}) };
    }));
  };
  const label = (point: Point | undefined): string => {
    if (!point || point.value.type === "blank" || point.state === "stale" || point.state === "unavailable") return "";
    if (point.value.type === "boolean") return accountText(point.value.value ? "TRUE" : "FALSE");
    if (point.value.type === "error") return accountText(point.value.value);
    const display = formatter.format(point.value.value, point.format);
    if (display.kind === "overflow") return accountText("####");
    if (display.parts.some(part => part.kind !== "text")) throw new DocumentError("UnsupportedFeature", "This chart label requires measured number-format spacing or repetition.");
    return accountText(display.parts.map(part => part.kind === "text" ? part.text : "").join(""));
  };
  const bind = (definition: DeepReadonly<WorkbookChart>, source: WorkbookChartData): WorkbookChartBinding => {
    const diagnostics: WorkbookChartDataDiagnostic[] = [], refs = new Map<string, WorkbookChartReferenceData>(), expected = new Map<string, DeepReadonly<WorkbookChartVector>>();
    const expectSlot = (role: Role, vector: DeepReadonly<WorkbookChartVector>, seriesId?: string): void => { expected.set(slotKey(role, seriesId), vector); };
    if (definition.categories) expectSlot("categories", definition.categories);
    if (definition.titleReference) expectSlot("title", definition.titleReference);
    if (definition.categoryAxisTitle?.reference) expectSlot("category-axis-title", definition.categoryAxisTitle.reference);
    if (definition.valueAxisTitle?.reference) expectSlot("value-axis-title", definition.valueAxisTitle.reference);
    for (const series of definition.series) { expectSlot("series-values", series.values, series.seriesId); if (series.name?.kind === "reference") expectSlot("series-name", series.name.reference, series.seriesId); }
    for (const ref of source.references) {
      const key = slotKey(ref.role, "seriesId" in ref ? ref.seriesId : undefined);
      if (!expected.has(key) || refs.has(key)) throw new DocumentError("InvalidDocument", "Chart data contains an unexpected or duplicate source slot.");
      const original = expected.get(key)!;
      if ("kind" in original) {
        const incoming = ref.data;
        if (incoming.kind !== "literal" || incoming.count !== original.count || incoming.valueType !== original.valueType || incoming.numberFormat !== original.numberFormat
          || incoming.points.length !== original.points.length || incoming.points.some((point, index) => point.index !== original.points[index]!.index || point.value !== original.points[index]!.value || point.numberFormat !== original.points[index]!.numberFormat))
          throw new DocumentError("InvalidDocument", "Chart literal data differs from its sole authoritative vector.");
        refs.set(key, original);
      } else {
        if (ref.data.kind === "literal") throw new DocumentError("InvalidDocument", "A chart reference cannot be replaced by a literal or cached vector.");
        refs.set(key, ref.data);
      }
    }
    if (refs.size !== expected.size) throw new DocumentError("InvalidDocument", "Chart data is missing a required source slot.");
    let unavailable = false;
    const visibleOnly = definition.visibleCellsOnly !== false;
    const read = (role: Role, seriesId?: string): readonly Point[] => {
      const data = refs.get(slotKey(role, seriesId)); if (!data) return [];
      const identity = { chartId: definition.chartId, role, ...(seriesId ? { seriesId } : {}) };
      if (data.kind === "error" || data.kind === "unresolved") {
        diagnostics.push(Object.freeze({ ...identity, code: data.kind === "error" ? "reference-error" : "unresolved-reference", detail: data.kind === "error" ? data.code : data.reason })); unavailable = true; return [];
      }
      const result = vector(data);
      for (const [index, point] of result.entries()) {
        if (visibleOnly && point.hidden) continue;
        if (point.state === "stale" || point.state === "unavailable") { unavailable = true; diagnostics.push(Object.freeze({ ...identity, index, code: point.state === "stale" ? "stale-cell" : "unavailable-cell", ...(point.issue ? { issue: point.issue } : {}) })); }
        if (point.state === "imported") diagnostics.push(Object.freeze({ ...identity, index, code: "imported-cell" }));
        if (role === "series-values" && point.value.type === "error") diagnostics.push(Object.freeze({ ...identity, index, code: "error-point", detail: point.value.value }));
      }
      return result;
    };
    const scalarLabel = (role: Role, seriesId?: string): string => {
      const values = read(role, seriesId);
      if (values.length > 1) {
        unavailable = true; diagnostics.push(Object.freeze({ chartId: definition.chartId, role, code: "non-scalar-label", ...(seriesId ? { seriesId } : {}) })); return "";
      }
      return visibleOnly && values[0]?.hidden ? "" : label(values[0]);
    };
    const categories = read("categories"), titles: Partial<Record<"title" | "category-axis-title" | "value-axis-title", string>> = {};
    for (const role of ["title", "category-axis-title", "value-axis-title"] as const) if (refs.has(slotKey(role))) {
      titles[role] = scalarLabel(role);
    }
    const series = definition.series.map((series, index) => ({ seriesId: series.seriesId,
      name: series.name?.kind === "text" ? accountText(series.name.text) : series.name?.kind === "reference"
        ? scalarLabel("series-name", series.seriesId) : accountText(`Series ${index + 1}`),
      points: read("series-values", series.seriesId) }));
    if (unavailable) return Object.freeze({ kind: "unavailable", chartId: definition.chartId, diagnostics: Object.freeze(diagnostics) });
    const count = Math.max(categories.length, ...series.map(series => series.points.length));
    // Padding short vectors is derived allocation too; charge the complete
    // rectangular plot before materializing its slots.
    accountPoints(count * (series.length + 1), true);
    const indexes = Array.from({ length: count }, (_, index) => index).filter(index => !visibleOnly
      || !series.some(series => series.points[index]) || series.some(series => series.points[index] && !series.points[index]!.hidden));
    const data: BasicChartData = Object.freeze({
      categories: Object.freeze(indexes.map(index => definition.categories ? visibleOnly && categories[index]?.hidden ? "" : label(categories[index]) : accountText(String(index + 1)))),
      sourceIndexes: Object.freeze(indexes),
      series: Object.freeze(series.map(series => Object.freeze({ seriesId: series.seriesId, name: series.name,
        hidden: visibleOnly && series.points.length > 0 && series.points.every(point => point.hidden),
        values: Object.freeze(indexes.map(index => {
          const point = series.points[index]; if (!point || visibleOnly && point.hidden) return null;
          return point.value.type === "number" ? point.value.value : point.value.type === "boolean" ? Number(point.value.value) : null;
        })),
        missing: Object.freeze(indexes.map(index => {
          const point = series.points[index]; return visibleOnly && point?.hidden ? "span" as const : point?.value.type === "error" ? "gap" as const : undefined;
        })),
      }))),
    });
    let valueAxisFormat: WorkbookDisplayFormat | undefined;
    axisFormat: for (const item of series) for (const point of item.points) {
      if (visibleOnly && point.hidden || point.value.type !== "number" && point.value.type !== "boolean") continue;
      valueAxisFormat = point.format; break axisFormat;
    }
    return Object.freeze({ kind: "available", chartId: definition.chartId, data, titles: Object.freeze(titles), diagnostics: Object.freeze(diagnostics),
      ...(valueAxisFormat ? { valueAxisFormat } : {}),
      categoryNumbers: Object.freeze(indexes.map(index => { const point = categories[index]; return definition.categories?.valueType === "number" && point?.value.type === "number" && !(visibleOnly && point.hidden) ? point.value.value : null; })),
      valueFormats: Object.freeze(series.map(series => Object.freeze(indexes.map(index => series.points[index]?.format)))) });
  };
  return Object.freeze(response.charts.map(source => bind(charts.get(source.chartId)!, source)));
}
