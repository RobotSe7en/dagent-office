import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import { writeOpcZip, readDocumentZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { WorkbookDrawings } from "../src/xlsx/drawing.js";
import { WorkbookContent } from "../src/xlsx/content.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
async function read(bytes: Uint8Array) {
  const records = await readXlsxRecords(bytes), content = new WorkbookContent(records.limits.source), drawings = new WorkbookDrawings(records.opc, records.styles, content);
  for (const sheet of records.sheets) drawings.sheet(sheet.source.sheetId, sheet.root);
  const files = new Map([...createWorkbook({ documentId: "drawings" }), ...records.chunks.files, ...drawings.files]);
  const manifest = JSON.parse(decoder.decode(files.get("manifest.json")!)); manifest.chunks = records.chunks.descriptors; manifest.assets = drawings.assets; manifest.content = "content.html";
  const metadata = { workbook: { dateSystem: records.native.dateSystem, calculationMode: records.native.calculationMode, activeSheetId: records.native.activeSheetId,
    sheets: records.sheets.map(sheet => sheet.metadata), names: records.native.names, themeCssClass: "xlsx-theme", numberFormats: records.styles.numberFormats(),
    objects: drawings.objects, objectHyperlinks: drawings.hyperlinks } };
  files.set("manifest.json", encoder.encode(JSON.stringify(manifest))); files.set("metadata.json", encoder.encode(JSON.stringify(metadata)));
  files.set("content.html", encoder.encode(content.html()!)); files.set("styles.css", encoder.encode(records.styles.stylesheet()));
  return { model: openDocument(files, { kind: "workbook" }), drawings };
}
const source = () => readFileSync("examples/xlsx/08-images-shapes.xlsx");
function variant(edit: (xml: string) => string, relationships?: string): Uint8Array {
  const files = new Map(readDocumentZip(source())), part = "xl/drawings/drawing1.xml", original = decoder.decode(files.get(part)!);
  const changed = edit(original); expect(changed).not.toBe(original); files.set(part, encoder.encode(changed));
  if (relationships) {
    const path = "xl/drawings/_rels/drawing1.xml.rels";
    files.set(path, encoder.encode(decoder.decode(files.get(path)!).replace('</Relationships>', relationships + '</Relationships>')));
  }
  return writeOpcZip(files);
}
describe("native worksheet images and text drawings", () => {
  it.each(["examples/xlsx/08-images-shapes.xlsx", "examples/xlsx/producer/08-images-shapes.xlsx"])("preserves %s as editable package source with exact text, assets and anchor modes", async path => {
    const { model, drawings } = await read(readFileSync(path)), snapshot = model.snapshot(), objects = snapshot.metadata.workbook.objects!;
    expect(objects).toHaveLength(3); expect(objects[0]?.kind).toBe("image"); expect(objects[1]).toMatchObject({ kind: "geometry", preset: "roundRect" });
    expect(model.getNode(objects[1]!.nodeId)?.text).toBe("随单元格移动缩放"); expect(model.getNode(objects[2]!.nodeId)?.text).toBe("绝对位置文本框 ABC 中文😀");
    const image = model.getNode(objects[0]!.nodeId)!;
    expect(image.attributes["data-office-rotation"]).toBe("15"); expect(image.attributes["data-office-width"]).toBeUndefined();
    expect(drawings.assets).toHaveLength(1); expect([...drawings.files.values()][0]).toEqual([...readDocumentZip(readFileSync(path))].find(([name]) => name.startsWith("xl/media/"))?.[1]);
    if (!path.includes("producer")) {
      expect(objects[0]?.anchor).toEqual({ kind: "one-cell", from: { row: 2, column: 1, offsetXPt: 0, offsetYPt: 0 }, widthPt: 160, heightPt: 120 });
      expect(objects[1]?.anchor).toMatchObject({ kind: "two-cell", moveWithCells: true, sizeWithCells: true });
      expect(objects[2]).toMatchObject({ kind: "text", anchor: { kind: "absolute", xPt: 60, yPt: 225, widthPt: 300, heightPt: 90 } });
      expect(image.attributes["data-office-crop-left"]).toBe("0.1");
    } else {
      expect(objects[0]?.anchor).toMatchObject({ kind: "two-cell", moveWithCells: true, sizeWithCells: false });
      expect(objects[2]).toMatchObject({ kind: "geometry", preset: "rect", anchor: { kind: "two-cell", moveWithCells: false, sizeWithCells: false } });
      expect(model.getNode(objects[1]!.nodeId)?.attributes["data-office-adjust-adj"]).toBe("16667");
    }
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(snapshot);
  });
  it("keeps background style references, text styles, locks and object/run hyperlinks with their owning source nodes", async () => {
    const style = '<xdr:style><a:lnRef idx="0"/><a:fillRef idx="1001"><a:schemeClr val="accent2"/></a:fillRef><a:effectRef idx="1"/><a:fontRef idx="minor"/></xdr:style>';
    const bytes = variant(xml => xml.replace('<xdr:txBody>', style + '<xdr:txBody>')
      .replace('<xdr:cNvPicPr/>', '<xdr:cNvPicPr preferRelativeResize="0"><a:picLocks noChangeAspect="1" noCrop="0"/></xdr:cNvPicPr>')
      .replace('descr="32 × 24 蓝色矩形"/>', 'descr="32 × 24 蓝色矩形"><a:hlinkClick r:id="web" tooltip="来源"/></xdr:cNvPr>')
      .replace('<a:rPr lang="zh-CN" sz="1400">', '<a:rPr lang="zh-CN" sz="1400"><a:hlinkClick r:id="internal"/>'),
    '<Relationship Id="web" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.test/source" TargetMode="External"/><Relationship Id="internal" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="#Sheet1!A1" TargetMode="External"/>');
    const { model } = await read(bytes), objects = model.snapshot().metadata.workbook.objects!;
    expect(objects[0]).toMatchObject({ locks: { noChangeAspect: true, noCrop: false }, preferRelativeResize: false });
    expect(model.getNode(objects[1]!.nodeId)?.attributes.class).toContain("xlsx-theme-background-1");
    expect(model.snapshot().metadata.workbook.objectHyperlinks).toEqual([
      { nodeId: objects[0]!.nodeId, url: "https://example.test/source", tooltip: "来源" },
      expect.objectContaining({ location: "Sheet1!A1" }),
    ]);
  });
  it("fails unsafe, ambiguous or unsupported source objects without returning a partial document", async () => {
    const edits = [
      (s: string) => s.replace('id="2" name="Shape 2"', 'id="1" name="Shape 2"'),
      (s: string) => s.replace('<xdr:sp>', '<xdr:sp macro="run" >'),
      (s: string) => s.replace('prst="roundRect"', 'prst="star5"'),
      (s: string) => s.replace('l="10000"', 'l="-10000"'),
      (s: string) => s.replace('r="10000"', 'r="90000"'),
      (s: string) => s.replace('r:embed="image"', 'r:embed="missing"'),
      (s: string) => s.replace('r:embed="image"', 'r:link="image"'),
      (s: string) => s.replace('<a:bodyPr wrap="square"/>', '<a:bodyPr wrap="square" unknown="1"/>'),
      (s: string) => s.replace('<a:rPr lang="zh-CN" sz="1400">', '<a:rPr lang="zh-CN" sz="1400" unknown="1">'),
    ];
    for (const [index, edit] of edits.entries()) await expect(read(variant(edit)), `mutation ${index}`).rejects.toBeDefined();
    const records = await readXlsxRecords(source()), content = new WorkbookContent({ ...records.limits.source, maxFileBytes: 300 }), drawings = new WorkbookDrawings(records.opc, records.styles, content);
    expect(() => drawings.sheet("sheet-1", records.sheets[0]!.root)).toThrowError(expect.objectContaining({ code: "ResourceLimitExceeded" }));
  });
  it("preserves styled explicit line breaks without losing UTF-16 text or paragraph-end properties", async () => {
    const { model } = await read(variant(xml => xml.replace('<a:endParaRPr lang="zh-CN"/>', '<a:br><a:rPr sz="1800" b="1"/></a:br><a:r><a:t xml:space="preserve"> 续行😀 </a:t></a:r><a:endParaRPr lang="zh-CN" sz="1000"/>')));
    const shape = model.snapshot().metadata.workbook.objects![1]!, html = model.getNode(shape.nodeId)!;
    expect(html.text).toBe("随单元格移动缩放\n 续行😀 "); expect(html.html).toMatch(/<br[^>]+class=/);
    expect(html.html).toContain("data-office-end-style");
    expect(openDocument(model.serialize(), { kind: "workbook" }).snapshot()).toEqual(model.snapshot());
  });
});
