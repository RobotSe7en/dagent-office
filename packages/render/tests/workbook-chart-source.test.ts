import { createHash } from "node:crypto";
import { readFileSync } from "node:fs";
import { execFileSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookChart, WorkbookDocumentMetadata, WorkbookManifest, WorkbookModel } from "@dagent-office/core";
import { readXlsxRecords } from "../../office/src/xlsx/read.js";
import { projectXlsxRecords } from "../../office/tests/helpers/xlsx-record-model.js";
import { DEFAULT_CHART_READ_LIMITS, modelChartSource } from "../src/workbook/chart-source.js";
import type { WorkbookChartReadRequest, WorkbookChartReadResponse } from "../src/workbook/chart-source.js";

const encode = (text: string) => new TextEncoder().encode(text), decode = (bytes: Uint8Array) => new TextDecoder().decode(bytes);
const reference = (expression: string) => ({ expression, valueType: "number" as const, cache: { state: "imported" as const, count: 1, points: [{ index: 0, value: 999 }] } });
function seeded(edit: (chart: WorkbookChart, metadata: WorkbookDocumentMetadata) => void = () => {}): WorkbookModel {
  const files = new Map(createWorkbook({ documentId: "chart-data", sheetName: "Data" }));
  const manifest = JSON.parse(decode(files.get("manifest.json")!)) as WorkbookManifest;
  const metadata = JSON.parse(decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets.push({ ...metadata.workbook.sheets[0]!, sheetId: "second", name: "第二表" });
  manifest.content = "content.html";
  metadata.workbook.objects = [{ kind: "chart", nodeId: "chart-node", chartId: "chart", sheetId: "sheet-1", locked: true, printable: true, anchor: { kind: "absolute", xPt: 0, yPt: 0, widthPt: 400, heightPt: 250 } }];
  const chart: WorkbookChart = { chartId: "chart", nodeId: "chart-node", kind: "column", legend: "bottom", showLabels: false,
    categories: { expression: "=A1:A2", valueType: "text" }, series: [{ seriesId: "values", values: reference("=B1:B6") }] };
  metadata.workbook.charts = [chart]; edit(chart, metadata);
  files.set("content.html", encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="content" data-office-kind="workbook-content"><div data-node-id="chart-node" data-office-kind="chart"></div></section></body></html>'));
  const path = "sheets/sheet-1/0-0.html", bytes = encode('<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody>'
    + '<tr data-row="0"><td data-column="0" data-value-type="text">中文😀</td><td data-column="1" data-value-type="number" data-value="12">旧显示</td></tr>'
    + '<tr data-row="1"><td data-column="0" data-value-type="boolean" data-value="true">TRUE</td><td data-column="1" data-value-type="number" data-value="23" data-formula="=20+3" data-cache-state="imported">23</td></tr>'
    + '<tr data-row="2"><td data-column="1" data-value-type="number" data-value="34" data-formula="=30+4" data-cache-state="valid">34</td></tr>'
    + '<tr data-row="3"><td data-column="1" data-value-type="number" data-value="45" data-formula="=40+5" data-cache-state="stale">45</td></tr>'
    + '<tr data-row="4"><td data-column="1" data-value-type="blank" data-formula="=UNSUPPORTED()" data-cache-state="unavailable"></td></tr>'
    + '<tr data-row="5"><td data-column="1" data-value-type="error" data-value="#N/A">#N/A</td></tr>'
    + '</tbody></table></body></html>');
  files.set(path, bytes); manifest.chunks.push({ sheetId: "sheet-1", row: 0, column: 0, path, sha256: createHash("sha256").update(bytes).digest("hex") });
  files.set("manifest.json", encode(JSON.stringify(manifest))); files.set("metadata.json", encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: "workbook" });
}
const request = (model: WorkbookModel, overrides: Partial<WorkbookChartReadRequest> = {}): WorkbookChartReadRequest => ({ ...DEFAULT_CHART_READ_LIMITS,
  documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision, chartIds: ["chart"], ...overrides });
const read = (model: WorkbookModel, overrides: Partial<WorkbookChartReadRequest> = {}, signal = new AbortController().signal) => modelChartSource(model).read(request(model, overrides), signal);
function cells(response: WorkbookChartReadResponse, role = "series-values") {
  const result = response.charts[0]!.references.find(item => item.role === role)!.data;
  if (result.kind !== "cells") throw new Error("Expected source cells");
  return result.areas.flatMap(area => area.cells);
}

describe("bounded current-source workbook chart reads", () => {
  it("returns immutable sparse literals without resolving formulas or reading worksheet cells", async () => {
    const model = seeded(chart => {
      chart.categories = { kind: "literal", valueType: "text", count: 4, points: [{ index: 0, value: "季度😀" }] };
      chart.series[0]!.values = { kind: "literal", valueType: "number", count: 4, numberFormat: "0.00", points: [{ index: 0, value: -3 }, { index: 2, value: null }, { index: 3, value: "#N/A", numberFormat: "0%" }] };
    });
    const provider = { ...model, resolveReference() { throw new Error("No literal has a reference"); }, readWindow() { throw new Error("No literal owns cells"); } };
    const response = await read(provider, { maxPoints: 8 });
    expect(response.charts[0]!.references.map(ref => ref.data)).toEqual([model.snapshot().metadata.workbook.charts![0]!.categories, model.snapshot().metadata.workbook.charts![0]!.series[0]!.values]);
    expect(structuredClone(response)).toEqual(response); expect(Object.isFrozen(response.charts[0]!.references[0]!.data)).toBe(true);
    await expect(read(provider, { maxPoints: 7 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
    await expect(read(provider, { maxCodeUnits: 24 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it("preflights literal counts before reading any referenced cells and yields during literal text admission", async () => {
    let queried = 0;
    const model = seeded(chart => { chart.series[0]!.values = { kind: "literal", valueType: "number", count: 1024, points: Array.from({ length: 1024 }, (_, index) => ({ index, value: index })) }; });
    const provider = { ...model, readWindow(window: Parameters<WorkbookModel["readWindow"]>[0]) { queried++; return model.readWindow(window); } };
    await expect(read(provider, { maxPoints: 1025 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" }); expect(queried).toBe(0);
    const controller = new AbortController(), pending = read(provider, {}, controller.signal); controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(queried).toBe(0);
    const response = await read(provider); expect(cells(response, "categories")).toHaveLength(2); expect(response.charts[0]!.references[1]!.data.kind).toBe("literal");
  });
  it("keeps typed values, formula states and sparse blanks distinct from imported chart caches", async () => {
    const model = seeded(), source = model.serialize(), response = await read(model);
    expect(cells(response).map(cell => cell.state)).toEqual(["constant", "imported", "valid", "stale", "unavailable", "constant"]);
    expect(cells(response).map(cell => cell.value)).toEqual([{ type: "number", value: 12 }, { type: "number", value: 23 }, { type: "number", value: 34 }, { type: "number", value: 45 }, { type: "blank" }, { type: "error", value: "#N/A" }]);
    expect(cells(response, "categories").map(cell => cell.value)).toEqual([{ type: "text", value: "中文😀" }, { type: "boolean", value: true }]);
    expect(model.snapshot().metadata.workbook.charts![0]!.series[0]!.values).toMatchObject({ cache: { points: [{ index: 0, value: 999 }] } });
    expect(model.serialize()).toEqual(source); expect(Object.isFrozen(response.charts[0]?.references)).toBe(true); expect(Object.isFrozen(cells(response)[0]?.value)).toBe(true);
    expect(structuredClone(response)).toEqual(response);
  });
  it("retains exact region order, repeated cells, cross-sheet blanks and rectangular row order", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=(B3:B4,第二表!XFD1048576,A1:B2,B1)"); });
    const response = await read(model), data = response.charts[0]!.references[0]!.data;
    expect(data.kind === "cells" && data.areas.map(area => area.range.sheetId)).toEqual(["sheet-1", "second", "sheet-1", "sheet-1"]);
    expect(cells(response).map(cell => [cell.row, cell.column])).toEqual([[2, 1], [3, 1], [1048575, 16383], [0, 0], [0, 1], [1, 0], [1, 1], [0, 1]]);
    expect(cells(response)[2]).toMatchObject({ state: "constant", value: { type: "blank" } });
    await expect(read(model, { maxPoints: 7 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it("resolves names in the owning worksheet rather than the active sheet, including title and series slots", async () => {
    const model = seeded((chart, metadata) => {
      metadata.workbook.activeSheetId = "second";
      metadata.workbook.names = [{ nameId: "local", name: "LocalValues", sheetId: "sheet-1", expression: "=$B$1:$B$2", origin: { row: 0, column: 0 } }];
      chart.series[0]!.values = reference("=LocalValues"); chart.series[0]!.name = { kind: "reference", reference: { expression: "=A1", valueType: "text" } };
      chart.titleReference = { expression: "=A1", valueType: "text" }; chart.categoryAxisTitle = { reference: { expression: "=A1", valueType: "text" } }; chart.valueAxisTitle = { reference: { expression: "=A2", valueType: "text" } };
    });
    const response = await read(model);
    expect(response.charts[0]!.references.map(ref => ref.role)).toEqual(["categories", "title", "category-axis-title", "value-axis-title", "series-name", "series-values"]);
    expect(cells(response).map(cell => cell.value)).toEqual([{ type: "number", value: 12 }, { type: "number", value: 23 }]);
    expect(cells(response, "title")[0]!.value).toEqual({ type: "text", value: "中文😀" });
    expect(response.charts[0]!.references.at(-1)).toMatchObject({ seriesId: "values" });
  });
  it("returns dynamic, circular and broken reference states without using their imported caches", async () => {
    for (const [expression, expected] of [["=INDEX(A1:B5,0,1)", { kind: "unresolved", reason: "calculation-required" }], ["=Loop", { kind: "unresolved", reason: "circular-name" }], ["=Missing!B1", { kind: "error", code: "#REF!" }], ["=MissingName", { kind: "error", code: "#NAME?" }]] as const) {
      const model = seeded((chart, metadata) => { delete chart.categories; chart.series[0]!.values = reference(expression); metadata.workbook.names = [{ nameId: "loop", name: "Loop", expression: "=Loop", origin: { row: 0, column: 0 } }]; });
      expect((await read(model)).charts[0]!.references[0]!.data).toEqual(expected);
    }
  });
  it("preserves hidden flags and inherited number formats without filtering or coercing points", async () => {
    const model = seeded((chart, metadata) => {
      delete chart.categories; chart.series[0]!.values = reference("=A1:B3");
      metadata.workbook.numberFormats.push({ formatId: "row-format", code: "0.00" }, { formatId: "column-format", code: "0%" });
      const sheet = metadata.workbook.sheets[0]!; sheet.defaultRowsHidden = true;
      sheet.rowDimensions = [{ from: 1, to: 1, hidden: false, numberFormatId: "row-format" }];
      sheet.columnDimensions = [{ from: 1, to: 1, unit: "characters", hidden: true, numberFormatId: "column-format" }];
    });
    const points = cells(await read(model));
    expect(points.map(point => point.rowHidden)).toEqual([true, true, false, false, true, true]);
    expect(points.map(point => point.columnHidden)).toEqual([false, true, false, true, false, true]);
    expect(points[1]!.numberFormatId).toBe("column-format"); expect(points[3]!.numberFormatId).toBe("row-format");
  });
  it("admits all references before reading any cells and rejects invalid identity or duplicate requests", async () => {
    let reads = 0;
    const source = seeded(chart => { chart.series[0]!.values = reference("=B:B"); }), model = { ...source, readWindow: (window: Parameters<WorkbookModel["readWindow"]>[0]) => { reads++; return source.readWindow(window); } };
    for (const override of [{}, { maxReferences: 1 }, { maxCharts: 0 }, { chartIds: ["chart", "chart"] }, { chartIds: ["missing"] }, { revision: 999 }, { documentId: "other" }]) await expect(read(model, override)).rejects.toBeDefined();
    expect(reads).toBe(0);
    expect((await read(model, { chartIds: [] })).charts).toEqual([]);
  });
  it("counts actual UTF-16 values and repeated occurrences against the text budget", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=(A1,A1)"); });
    // chart id, owner, expression, series id, two area sheet ids, two four-unit values.
    const units = "chart".length + "sheet-1".length + "=(A1,A1)".length + "values".length + 2 * "sheet-1".length + 2 * "中文😀".length;
    expect(cells(await read(model, { maxCodeUnits: units }))).toHaveLength(2);
    await expect(read(model, { maxCodeUnits: units - 1 })).rejects.toMatchObject({ code: "ResourceLimitExceeded" });
  });
  it.each(["=A1:A1000", "=A1:ALL1", "=A1:AD20"])("yields bounded queries and cancels %s before reading the next batch", async expression => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference(expression); }), queried: number[] = [];
    const controller = new AbortController(), provider = modelChartSource({ ...model, readWindow(window) { queried.push(window.rows * window.columns); return model.readWindow(window); } });
    const pending = provider.read(request(model), controller.signal);
    expect(queried.length).toBeGreaterThan(0); expect(Math.max(...queried)).toBeLessThanOrEqual(256); const completed = [...queried]; controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(queried).toEqual(completed);
    await expect(provider.read(request(model), controller.signal)).rejects.toMatchObject({ code: "OperationCancelled" });
  });
  it("isolates requests before yielding and returns only the selected immutable revision", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series[0]!.values = reference("=A1:A512"); });
    const ids = ["chart"], input = { ...request(model), chartIds: ids }, pending = modelChartSource(model).read(input, new AbortController().signal);
    ids[0] = "missing"; input.documentId = "other"; input.revision = 999; input.maxPoints = 1;
    const response = await pending; expect(response.documentId).toBe("chart-data"); expect(response.revision).toBe(model.snapshot().manifest.revision); expect(cells(response)).toHaveLength(512);
  });
  it("allows cancellation during reference admission before any cells are read", async () => {
    const model = seeded(chart => { delete chart.categories; chart.series = Array.from({ length: 64 }, (_, index) => ({ seriesId: `series-${index}`, values: reference("=B1") })); });
    let bound = 0, queried = 0;
    const provider = modelChartSource({ ...model, resolveReference(expression, origin) { bound++; return model.resolveReference(expression, origin); }, readWindow(window) { queried++; return model.readWindow(window); } });
    const controller = new AbortController(), pending = provider.read(request(model), controller.signal);
    expect(bound).toBe(32); expect(queried).toBe(0); controller.abort();
    await expect(pending).rejects.toMatchObject({ code: "OperationCancelled" }); expect(bound).toBe(32); expect(queried).toBe(0);
  });
  it.each(["examples/xlsx/07-three-charts.xlsx", "examples/xlsx/producer/07-three-charts.xlsx"])("reads all three independent native chart sources from %s", async path => {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(path))), definitions = model.snapshot().metadata.workbook.charts!;
    const response = await read(model, { chartIds: definitions.map(chart => chart.chartId) });
    expect(response.charts).toHaveLength(3);
    for (const chart of response.charts) {
      const values = chart.references.filter(ref => ref.role === "series-values");
      expect(values[0]!.data.kind === "cells" && values[0]!.data.areas[0]!.cells.map(cell => cell.value)).toEqual([2, 5, 8].map(value => ({ type: "number", value })));
      expect(values.every(ref => ref.data.kind === "cells" && ref.data.areas[0]!.cells.every(cell => cell.state === "constant"))).toBe(true);
    }
  });
  it("reads mixed values and hidden-source coordinates independently of deliberately stale native chart caches", async () => {
    const bytes = execFileSync('python3', ['-c', "import importlib.util,sys; s=importlib.util.spec_from_file_location('chart_values','scripts/check-xlsx-chart-values.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))"], { maxBuffer: 4 * 1024 * 1024 });
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-chart-value-semantics.json', 'utf8')) as { sourceSha256: string; cases: { hidden: string; visibleOnly: boolean }[] };
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const { model } = projectXlsxRecords(await readXlsxRecords(bytes)), definitions = model.snapshot().metadata.workbook.charts!;
    const response = await read(model, { chartIds: definitions.map(chart => chart.chartId) }); expect(response.charts).toHaveLength(15);
    for (const [index, chart] of response.charts.entries()) {
      const points = cells({ ...response, charts: [chart] });
      expect(points.map(point => point.value)).toEqual([{ type: 'number', value: 2 }, { type: 'text', value: '7' }, { type: 'text', value: 'word' }, { type: 'boolean', value: true }, { type: 'boolean', value: false }, { type: 'blank' }, { type: 'blank' }, { type: 'text', value: '' }, { type: 'error', value: '#N/A' }, { type: 'error', value: '#DIV/0!' }, { type: 'number', value: -3 }, { type: 'number', value: 0 }]);
      expect(points[7]!.state).toBe('imported');
      expect(points.filter(point => point.rowHidden).map(point => point.row)).toEqual(evidence.cases[index]!.hidden === 'rows' ? [2, 4, 11] : []);
      expect(points.every(point => point.columnHidden)).toBe(evidence.cases[index]!.hidden === 'values-column');
      expect(definitions[index]!.series[0]!.values).toMatchObject({ cache: { points: Array.from({ length: 12 }, (_, index) => ({ index, value: 999 })) } });
    }
  });
});
