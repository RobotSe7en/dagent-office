import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openDocument } from "@dagent-office/core";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder(), path = "examples/xlsx/07-three-charts.xlsx";
async function read(bytes: Uint8Array) { return projectXlsxRecords(await readXlsxRecords(bytes)); }
function variant(edit: (source: string) => string, part = "xl/charts/chart1.xml"): Uint8Array {
  const files = new Map(readDocumentZip(readFileSync(path))), source = decoder.decode(files.get(part)!), changed = edit(source);
  expect(changed).not.toBe(source); files.set(part, encoder.encode(changed)); return writeOpcZip(files);
}
describe("worksheet chart object and style mapping", () => {
  it.each([path, "examples/xlsx/producer/07-three-charts.xlsx"])("reopens three chart types from %s with current-source references and imported caches", async path => {
    const { model } = await read(readFileSync(path)), snapshot = model.snapshot(), charts = snapshot.metadata.workbook.charts!;
    expect(charts.map(chart => chart.kind)).toEqual(["column", "line", "pie"]); expect(snapshot.metadata.workbook.objects?.map(object => object.kind)).toEqual(["chart", "chart", "chart"]);
    expect(charts.map(chart => chart.series.length)).toEqual([2, 2, 1]); expect(charts.every(chart => chart.visibleCellsOnly)).toBe(true);
    expect(charts[0]?.categories?.expression).toBe("=Data!$A$2:$A$4");
    expect(charts[0]?.series[0]?.name).toMatchObject({ kind: "reference", reference: { expression: "=Data!$B$1" } });
    expect(charts[0]?.series[0]?.values).toMatchObject({ expression: "=Data!$B$2:$B$4", valueType: "number", cache: { state: "imported" } });
    for (const chart of charts) { expect(model.getNode(chart.titleNodeId!)?.text).toBe("图表 " + chart.kind); expect(chart.legend).toBe("bottom"); }
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  });
  it("preserves cell-linked titles and decoded number formats without materializing cached title text", async () => {
    const { model } = await read(variant(source => source.replace(/<c:tx><c:rich>[\s\S]*?<\/c:rich><\/c:tx>/, '<c:tx><c:strRef><c:f>Data!$A$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>旧标题</c:v></c:pt></c:strCache></c:strRef></c:tx>')
      .replace('formatCode="0" sourceLinked="0"', 'formatCode="0_x0025_" sourceLinked="0"')));
    const chart = model.snapshot().metadata.workbook.charts![0]!;
    expect(chart.titleNodeId).toBeUndefined(); expect(chart.titleCssClass).toBeDefined(); expect(chart.titleReference).toMatchObject({ expression: "=Data!$A$1", cache: { state: "imported" } });
    expect(chart.numberFormat).toBe("0%"); expect(model.getNode(chart.nodeId)?.text).toBe("");
    expect(model.snapshot().styles[0]?.css).toContain('--office-axis-number-format-source-linked:false');
  });
  it("retains the selected modern leader-line flag and rejects conflicting or unknown extensions", async () => {
    const extension = '<c:extLst><c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}"><c15:showLeaderLines xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart" val="1"/></c:ext></c:extLst>';
    const inject = (xml: string, extra: string) => xml.replace('<c:axId val="100"/>', '<c:dLbls><c:showVal val="1"/>' + extra + '</c:dLbls><c:axId val="100"/>');
    const { model } = await read(variant(xml => inject(xml, extension)));
    expect(model.snapshot().styles[0]?.css).toContain('--office-label-show-leader-lines:true');
    for (const extra of [extension + extension, '<c:showLeaderLines val="0"/>' + extension, extension.replace('CE6537A1', '00000000'), extension.replace('c15:showLeaderLines', 'c15:unknown'), extension.replace('val="1"', 'val="maybe"'), extension.replace('val="1"', 'val="1" unknown="1"')]) {
      await expect(read(variant(xml => inject(xml, extra)))).rejects.toBeDefined();
    }
  });
  it("normalizes native percentage amounts and present-element schema defaults", async () => {
    const { model } = await read(variant(xml => xml.replace('<c:barDir val="col"/>', '<c:barDir/>').replace('<c:grouping val="clustered"/>', '<c:grouping/>')
      .replace('<c:axId val="100"/>', '<c:gapWidth val="175%"/><c:overlap/><c:axId val="100"/>').replace('<c:legendPos val="b"/>', '<c:legendPos/>')
      .replace('<c:dispBlanksAs val="gap"/>', '<c:dispBlanksAs/>')));
    expect(model.snapshot().metadata.workbook.charts![0]).toMatchObject({ gapWidth: 175, overlap: 0, legend: "right", blankAs: "zero" });
  });
  it("keeps independent rich and cell-linked axis titles inside their owning chart", async () => {
    const { model } = await read(variant(xml => xml.replace('<c:catAx>', '<c:catAx><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>季度😀</a:t></a:r></a:p></c:rich></c:tx></c:title>')
      .replace('<c:valAx>', '<c:valAx><c:title><c:tx><c:strRef><c:f>Data!$B$1</c:f></c:strRef></c:tx></c:title>')));
    const chart = model.snapshot().metadata.workbook.charts![0]!;
    expect(model.getNode(chart.categoryAxisTitle!.nodeId!)?.text).toBe('季度😀');
    expect(chart.valueAxisTitle?.reference).toEqual({ expression: '=Data!$B$1', valueType: 'text' });
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
  });
  it("distinguishes absent flags from present booleans with schema-default true", async () => {
    const { model } = await read(variant(xml => xml.replace('<c:plotVisOnly val="1"/>', '<c:showDLblsOverMax/>')));
    const chart = model.snapshot().metadata.workbook.charts![0]!;
    expect(chart.visibleCellsOnly).toBeUndefined(); expect(chart.showLabelsOverMaximum).toBe(true);
  });
  it("rejects unsupported combinations, broken axes, duplicate points and silent unknown properties", async () => {
    const edits = [
      (s: string) => s.replace('val="clustered"', 'val="stacked"'),
      (s: string) => s.replace('<c:chartSpace ', '<c:chartSpace unknown="1" '),
      (s: string) => s.replace('<c:barDir val="col"/>', '<c:barDir val="col" unknown="1"/>'),
      (s: string) => s.replace('<c:crossAx val="200"/>', '<c:crossAx val="100"/>'),
      (s: string) => s.replace('<c:axId val="200"/>', '<c:axId val="300"/>'),
      (s: string) => s.replace('Data!$C$2:$C$4', '[External.xlsx]Data!$C$2:$C$4'),
      (s: string) => s.replace('<c:ser>', '<c:ser><c:extLst/>'),
      (s: string) => s.replace('<c:plotArea>', '<c:plotArea><c:bar3DChart/>'),
      (s: string) => s.replace('<c:overlay val="0"/>', '<c:overlay val="1"/>'),
      (s: string) => s.replace('<c:catAx>', '<c:catAx><c:title/>'),
      (s: string) => s.replace('<c:scaling>', '<c:scaling><c:min val="0x10"/>'),
      (s: string) => s.replace('<c:axId val="100"/>', '<c:gapWidth val="3.5"/><c:axId val="100"/>'),
      (s: string) => s.replace('<c:axId val="100"/>', '<c:gapWidth val="501%"/><c:axId val="100"/>'),
      (s: string) => s.replace('<c:valAx>', '<c:valAx><c:minorUnit val="0"/>'),
      (s: string) => s.replace('<c:ser>', '<c:ser><c:marker><c:symbol val="circle"/></c:marker>'),
      (s: string) => s.replace('<c:ser>', '<c:ser><c:dPt><c:idx val="0"/></c:dPt><c:dPt><c:idx val="0"/></c:dPt>'),
    ];
    for (const [index, edit] of edits.entries()) await expect(read(variant(edit)), `mutation ${index}`).rejects.toBeDefined();
  });
});
