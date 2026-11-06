import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { DEFAULT_DOCUMENT_READ_LIMITS } from "@dagent-office/core";
import { openOpc } from "../src/ooxml/package.js";
import { C, elements, xml } from "../src/ooxml/xml.js";
import { readChartLabel, readChartReference, readChartVector } from "../src/xlsx/chart-data.js";
import type { ChartDataBudget } from "../src/xlsx/chart-data.js";

const encoder = new TextEncoder();
const node = (source: string) => xml(encoder.encode(source), "chart.xml", DEFAULT_DOCUMENT_READ_LIMITS, { nodes: 0 });
const budget = (): ChartDataBudget => ({ items: 0, bytes: 0, maxItems: 1000, maxBytes: 1_000_000 });
describe("native worksheet chart references and imported caches", () => {
  it("reads all six independent/producer charts while retaining worksheet formulas instead of creating constant data sets", () => {
    const total = budget(); let references = 0;
    for (const path of ["examples/xlsx/07-three-charts.xlsx", "examples/xlsx/producer/07-three-charts.xlsx"]) {
      const opc = openOpc(readFileSync(path), {});
      for (const part of opc.files.keys()) if (/^xl\/charts\/chart\d+\.xml$/.test(part)) {
        const visit = (source: ReturnType<typeof node>): void => {
          if (["strRef", "numRef"].includes(source.name)) {
            const result = readChartReference(source, total); references++;
            expect(result.expression).toMatch(/^=Data!\$[A-C]\$[1-4](?::\$[A-C]\$[1-4])?$/);
            expect(result.cache?.state).toBe("imported"); expect(result.cache?.count).toBe(result.expression.includes(":") ? 3 : 1);
          } else for (const child of elements(source, undefined, C)) visit(child);
        };
        visit(opc.read(part));
      }
    }
    expect(references).toBe(30); expect(total.items).toBe(100);
  });
  it("preserves missing caches, sparse points, per-point formats and UTF-16 escape semantics", () => {
    const absent = readChartReference(node(`<c:numRef xmlns:c="${C}"><c:f>_x0041_</c:f></c:numRef>`), budget());
    expect(absent).toEqual({ expression: "=_x0041_", valueType: "number" });
    const sparse = readChartReference(node(`<c:numRef xmlns:c="${C}"><c:f>Sheet1!B2:B5</c:f><c:numCache><c:formatCode>0.00</c:formatCode><c:ptCount val="4"/><c:pt idx="3" formatCode="0%"><c:v>0.5</c:v></c:pt><c:pt idx="0"><c:v>#N/A</c:v></c:pt><c:pt idx="1"><c:v/></c:pt></c:numCache></c:numRef>`), budget());
    expect(sparse.cache).toEqual({ state: "imported", count: 4, numberFormat: "0.00", points: [{ index: 0, value: "#N/A" }, { index: 1, value: null }, { index: 3, value: 0.5, numberFormat: "0%" }] });
    const text = readChartLabel(node(`<c:tx xmlns:c="${C}"><c:v>中文_xD83D__xDE00__x005F_x0041_</c:v></c:tx>`), budget());
    expect(text).toEqual({ kind: "text", text: "中文😀_x0041_" });
    const inferred = readChartReference(node(`<c:strRef xmlns:c="${C}"><c:f>A1:A3</c:f><c:strCache><c:pt idx="2"><c:v>末项</c:v></c:pt></c:strCache></c:strRef>`), budget());
    expect(inferred.cache?.count).toBe(3);
  });
  it("rejects malformed or external references, duplicate cache points, unknown metadata and ambiguous vectors", () => {
    for (const content of [
      '<c:f>[External.xlsx]Sheet1!A1</c:f>', '<c:f>=A1</c:f>', '<c:f>SUM(A1:A3)</c:f>',
      '<c:f>A1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="1"><c:v>2</c:v></c:pt></c:numCache>',
      '<c:f>A1:A2</c:f><c:numCache><c:ptCount val="2"/><c:pt idx="0"><c:v>2</c:v></c:pt><c:pt idx="0"><c:v>3</c:v></c:pt></c:numCache>',
      '<c:f>A1</c:f><c:numCache><c:ptCount val="1"/><c:pt idx="0"><c:v>oops</c:v></c:pt></c:numCache>',
      '<c:f>A1</c:f><c:extLst/>',
    ]) expect(() => readChartReference(node(`<c:numRef xmlns:c="${C}">${content}</c:numRef>`), budget())).toThrow();
    const literal = '<c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit>';
    expect(readChartVector(node(`<c:val xmlns:c="${C}">${literal}</c:val>`), budget(), true)).toEqual({ kind: "literal", valueType: "number", count: 1, points: [{ index: 0, value: 1 }] });
    expect(() => readChartVector(node(`<c:val xmlns:c="${C}">${literal}<c:numRef><c:f>A1</c:f></c:numRef></c:val>`), budget(), true)).toThrowError(expect.objectContaining({ code: "InvalidDocument" }));
  });
  it("enforces cumulative declared-point and metadata-byte budgets before expanding cache records", () => {
    const source = node(`<c:numRef xmlns:c="${C}"><c:f>A1:A3</c:f><c:numCache><c:ptCount val="3"/></c:numCache></c:numRef>`), count = { ...budget(), maxItems: 7 };
    readChartReference(source, count); expect(() => readChartReference(source, count)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    expect(() => readChartReference(source, { ...budget(), maxBytes: 100 })).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
    const huge = node(`<c:numRef xmlns:c="${C}"><c:f>A:A</c:f><c:numCache><c:ptCount val="4294967295"/></c:numCache></c:numRef>`);
    expect(() => readChartReference(huge, budget())).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
