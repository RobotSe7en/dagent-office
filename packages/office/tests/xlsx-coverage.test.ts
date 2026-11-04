import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const edit = (files: Map<string, Uint8Array>, path: string, before: string, after: string) => files.set(path, encoder.encode(decoder.decode(files.get(path)!).replace(before, after)));
async function read(files: Map<string, Uint8Array>) { return projectXlsxRecords(await readXlsxRecords(writeOpcZip(files))); }
describe("native source-part consumption", () => {
  it("rejects reachable but unreferenced worksheets, tables and charts", async () => {
    for (const [fixture, source, target, relationships, type, contentType] of [
      ['01-values-text', 'xl/worksheets/sheet1.xml', 'xl/worksheets/extra.xml', 'xl/_rels/workbook.xml.rels', 'worksheet', 'spreadsheetml.worksheet'],
      ['05-native-table', 'xl/tables/table1.xml', 'xl/tables/extra.xml', 'xl/worksheets/_rels/sheet1.xml.rels', 'table', 'spreadsheetml.table'],
      ['07-three-charts', 'xl/charts/chart1.xml', 'xl/charts/extra.xml', 'xl/drawings/_rels/drawing1.xml.rels', 'chart', 'drawingml.chart'],
    ]) {
      const files = new Map(readDocumentZip(readFileSync(`examples/xlsx/${fixture}.xlsx`)));
      expect(files.has(source!)).toBe(true); files.set(target!, files.get(source!)!);
      edit(files, '[Content_Types].xml', '</Types>', `<Override PartName="/${target}" ContentType="application/vnd.openxmlformats-officedocument.${contentType}+xml"/></Types>`);
      edit(files, relationships!, '</Relationships>', `<Relationship Id="extra" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/${type}" Target="/${target}"/></Relationships>`);
      await expect(read(files), fixture).rejects.toMatchObject({ code: 'InvalidDocument', message: 'A native workbook part has no consumed source feature.' });
    }
  });
  it("validates a package thumbnail before discarding the replaceable preview", async () => {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/08-images-shapes.xlsx'))), image = [...files.keys()].find(path => path.endsWith('.png'))!;
    files.set('docProps/thumbnail.png', files.get(image)!);
    edit(files, '[Content_Types].xml', '</Types>', '<Override PartName="/docProps/thumbnail.png" ContentType="image/png"/></Types>');
    edit(files, '_rels/.rels', '</Relationships>', '<Relationship Id="thumbnail" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" Target="docProps/thumbnail.png"/></Relationships>');
    const { model } = await read(files); expect(model.snapshot().manifest.assets).toHaveLength(1);
    files.set('docProps/thumbnail.png', encoder.encode('<script/>')); await expect(read(files)).rejects.toBeDefined();
  });
});
