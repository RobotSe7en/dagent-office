import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "../src/index.js";
import type { WorkbookChart, WorkbookDocumentMetadata, WorkbookManifest } from "../src/index.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function seeded(edit?: (chart: WorkbookChart, metadata: WorkbookDocumentMetadata) => void) {
  const files = new Map(createWorkbook({ documentId: "charts" })), manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)) as WorkbookManifest;
  const metadata = JSON.parse(decoder.decode(files.get("metadata.json")!)) as WorkbookDocumentMetadata;
  manifest.content = "content.html";
  metadata.workbook.objects = [{ kind: "chart", nodeId: "chart-node", chartId: "chart", sheetId: "sheet-1", locked: true, printable: true, anchor: { kind: "absolute", xPt: 0, yPt: 0, widthPt: 400, heightPt: 250 } }];
  const chart: WorkbookChart = { chartId: "chart", nodeId: "chart-node", kind: "column", legend: "bottom", showLabels: true, labelContents: ["value"], titleNodeId: "title", visibleCellsOnly: true, varyColors: false, showLabelsOverMaximum: false,
    categories: { expression: "=A2:A4", valueType: "text", cache: { state: "imported", count: 3, points: [{ index: 0, value: "中文😀" }, { index: 2, value: "尾项" }] } },
    series: [{ seriesId: "series-0", nativeIndex: 7, name: { kind: "reference", reference: { expression: "=B1", valueType: "text" } }, values: { expression: "=B2:B4", valueType: "number", cache: { state: "imported", count: 3, points: [{ index: 0, value: 12 }, { index: 1, value: "#N/A" }, { index: 2, value: null }], numberFormat: "0.00" } } }],
    seriesStyles: [{ seriesId: "series-0", cssClass: "cell-default" }],
  };
  metadata.workbook.charts = [chart]; edit?.(chart, metadata);
  files.set("content.html", encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="styles.css"></head><body><section data-node-id="content" data-office-kind="workbook-content"><div data-node-id="chart-node" data-office-kind="chart"><div data-node-id="title"><p data-node-id="title-p">当前来源</p></div></div></section></body></html>'));
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); files.set("metadata.json", encoder.encode(JSON.stringify(metadata))); return files;
}
describe("workbook chart source references and derived caches", () => {
  it("roundtrips independent references, sparse imported caches and rich source titles", () => {
    const model = openDocument(seeded(), { kind: "workbook" }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.charts?.[0]?.series[0]?.values.expression).toBe("=B2:B4");
    expect(snapshot.metadata.workbook.charts?.[0]?.categories?.cache?.points).toHaveLength(2);
    expect(model.getNode("title")?.text).toBe("当前来源"); expect(model.getCell("sheet-1", 1, 1)).toBeUndefined();
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  });
  it("accepts named, structured, exact multi-region and invalidated references without inventing cache values", () => {
    for (const expression of ["=Amounts", "=Sales[Revenue]", "=(A1:A2,A4:A5,C3)", "=#REF!", "=Sheet1!A:A", "=INDEX(A1:B5,0,1)"]) {
      const model = openDocument(seeded(chart => { chart.series[0]!.values = { expression, valueType: "number" }; }), { kind: "workbook" });
      expect(model.snapshot().metadata.workbook.charts?.[0]?.series[0]?.values).toEqual({ expression, valueType: "number" });
    }
  });
  it("rejects competing title sources, orphan objects, invalid cache points and ambiguous style ownership", () => {
    const edits: NonNullable<Parameters<typeof seeded>[0]>[] = [
      chart => { chart.series[0]!.values.expression = "=SUM(A1:A3)"; },
      chart => { chart.titleReference = { expression: "=A1", valueType: "text" }; },
      chart => { chart.titleNodeId = "title-p"; },
      chart => { chart.categoryAxisTitle = { nodeId: "title" }; },
      chart => { chart.valueAxisTitle = {}; },
      chart => { chart.categoryAxisTitle = { reference: { expression: "=A1", valueType: "text" }, cssClass: "missing" }; },
      chart => { chart.series[0]!.values.cache!.points[1]!.index = 3; },
      chart => { chart.series[0]!.values.cache!.points[1]!.index = 0; },
      chart => { chart.series[0]!.values.cache!.points[1]!.value = "not numeric"; },
      chart => { chart.categories!.cache!.numberFormat = "0"; },
      chart => { chart.seriesStyles![0]!.seriesId = "missing"; },
      chart => { chart.plotCssClass = "missing"; },
      chart => { chart.series.push(chart.series[0]!); },
      chart => { chart.minimum = 10; chart.maximum = 5; },
      (chart, metadata) => { metadata.workbook.charts!.push(chart); },
      (_, metadata) => { metadata.workbook.charts = []; },
    ];
    for (const edit of edits) expect(() => openDocument(seeded(edit))).toThrow();
    expect(() => openDocument(seeded(chart => { chart.categories!.cache!.count = 0xffff_ffff; }))).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
});
