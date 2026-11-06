import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openDocument } from "@doffi/core";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
function source(core?: (s: string) => string, app?: (s: string) => string): Uint8Array {
  const files = new Map(readDocumentZip(readFileSync("examples/xlsx/producer/01-values-text.xlsx")));
  for (const [path, edit] of [["docProps/core.xml", core], ["docProps/app.xml", app]] as const) if (edit) {
    const original = decoder.decode(files.get(path)!), next = edit(original); expect(next).not.toBe(original); files.set(path, encoder.encode(next));
  }
  return writeOpcZip(files);
}
const vectors = '<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs><TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Values</vt:lpstr></vt:vector></TitlesOfParts>';
describe("workbook document properties", () => {
  it("keeps title, locale and file provenance through the unified document package", async () => {
    const records = await readXlsxRecords(source(xml => xml.replace('<dc:title></dc:title>', '<dc:title>季度_x0041_😀</dc:title>').replace('<dc:creator></dc:creator>', '<dc:creator>文档作者</dc:creator>')
      .replace('</cp:coreProperties>', '<dcterms:created xsi:type="dcterms:W3CDTF">2026-11-04T11:22:33+08:00</dcterms:created></cp:coreProperties>'), xml => xml.replace('</Properties>', '<Company>示例公司</Company>' + vectors + '</Properties>')));
    expect(records.properties).toMatchObject({ title: '季度_x0041_😀', locale: 'en-US', properties: { creator: '文档作者', company: '示例公司', created: '2026-11-04T11:22:33+08:00', totalEditingMinutes: 0 } });
    const model = projectXlsxRecords(records).model;
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
    expect(model.snapshot().metadata.properties?.creator).toBe('文档作者');
  });
  it("rejects unknown or duplicate core fields, nested markup and misleading date types", async () => {
    for (const edit of [
      (s: string) => s.replace('</cp:coreProperties>', '<dc:title>duplicate</dc:title></cp:coreProperties>'),
      (s: string) => s.replace('<dc:creator></dc:creator>', '<dc:creator unknown="1"/>'),
      (s: string) => s.replace('<dc:creator></dc:creator>', '<dc:creator><dc:title>nested</dc:title></dc:creator>'),
      (s: string) => s.replace('</cp:coreProperties>', '<cp:unknown>value</cp:unknown></cp:coreProperties>'),
      (s: string) => s.replace('</cp:coreProperties>', '<dcterms:created xsi:type="dc:W3CDTF">2026-11-04</dcterms:created></cp:coreProperties>'),
    ]) await expect(readXlsxRecords(source(edit))).rejects.toBeDefined();
  });
  it("bounds application statistics and refuses unsupported active metadata", async () => {
    for (const xml of [vectors.replace('size="2"', 'size="3"'), vectors.replace('<vt:i4>1</vt:i4>', '<vt:lpstr>1</vt:lpstr>'), '<HyperlinkBase>https://example.invalid/</HyperlinkBase>', '<DigSig/>', '<ScaleCrop>sometimes</ScaleCrop>', '<Pages>-1</Pages>', '<Application>duplicate</Application>']) {
      await expect(readXlsxRecords(source(undefined, app => app.replace('</Properties>', xml + '</Properties>')))).rejects.toBeDefined();
    }
  });
});
