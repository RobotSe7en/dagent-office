import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";
import { readFileSync } from "node:fs";
import { createHash } from "node:crypto";
import { execFileSync } from "node:child_process";
import { describe, expect, it } from "vitest";
import { openDocument } from "@dagent-office/core";
import { importXlsx } from "../src/xlsx-import.js";
import { writeOpcZip, readDocumentZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { WorkbookDrawings } from "../src/xlsx/drawing.js";
import { WorkbookContent } from "../src/xlsx/content.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
async function read(bytes: Uint8Array) { return projectXlsxRecords(await readXlsxRecords(bytes)); }
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
  it("imports all independent text overflow modes without changing source text on reopen", async () => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-text-overflow.xlsx');
    const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('overflow','scripts/check-xlsx-text-overflow.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(regenerated).toEqual(bytes);
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-text-overflow.json', 'utf8')) as { sourceSha256: string; cases: { horizontal: string; vertical: string; wrap: string; paragraphs: string[] }[] };
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const model = openDocument(await importXlsx(bytes, { documentId: 'text-overflow' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.objects).toHaveLength(24);
    for (const [index, object] of snapshot.metadata.workbook.objects!.entries()) {
      const source = model.getNode(object.nodeId)!, expected = evidence.cases[index]!;
      expect(source.text).toBe(expected.paragraphs.join(''));
      expect(source.attributes.style).toContain('--office-text-horizontal-overflow:' + expected.horizontal);
      expect(source.attributes.style).toContain('--office-text-vertical-overflow:' + expected.vertical);
      expect(object).toMatchObject({ text: { wrap: expected.wrap === 'square', autoFit: 'none' } });
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it.each(['horzOverflow="ellipsis"', 'vertOverflow="hidden"', 'horzOverflow="CLIP"', 'vertOverflow=""'])("rejects invalid native overflow %s atomically", async attribute => {
    await expect(read(variant(xml => xml.replace('<a:bodyPr wrap="square"/>', `<a:bodyPr wrap="square" ${attribute}/>`)))).rejects.toMatchObject({ code: 'InvalidDocument' });
  });
  it("retains all independent text anchoring cases and their source styles on package reopen", async () => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-text-anchoring.xlsx');
    const regenerated = execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('anchoring','scripts/check-xlsx-text-anchoring.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"]);
    expect(regenerated).toEqual(bytes);
    const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-text-anchoring.json', 'utf8')) as { sourceSha256: string; cases: { paragraphs: string[] }[] };
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const { model } = await read(bytes), snapshot = model.snapshot(), objects = snapshot.metadata.workbook.objects!;
    expect(objects).toHaveLength(evidence.cases.length);
    for (const [index, object] of objects.entries()) expect(model.getNode(object.nodeId)?.text).toBe(evidence.cases[index]!.paragraphs.join(''));
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
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
