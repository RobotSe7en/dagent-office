import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openDocument } from "@dagent-office/core";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const root = '{33333333-3333-4333-8333-333333333333}', reply = '{44444444-4444-4444-8444-444444444444}', second = '{55555555-5555-4555-8555-555555555555}';
function source(edit: (xml: string) => string = value => value): Uint8Array {
  const files = new Map(readDocumentZip(readFileSync('examples/xlsx/10-modern-threads.xlsx')));
  const update = (path: string, before: string, after: string) => files.set(path, encoder.encode(decoder.decode(files.get(path)!).replace(before, after)));
  update('[Content_Types].xml', '</Types>', '<Override PartName="/xl/comments1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"/></Types>');
  update('xl/worksheets/_rels/sheet1.xml.rels', '</Relationships>', '<Relationship Id="legacy" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="../comments1.xml"/></Relationships>');
  files.set('xl/comments1.xml', encoder.encode(edit('<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>'
    + `<author>tc=${root}</author><author>兼容作者 tc=${second}</author><author>经典作者</author></authors><commentList>`
    + '<comment ref="B2" authorId="0"><text><t>只供旧应用使用的缓存</t></text></comment>'
    + '<comment ref="C3" authorId="1"><text><t>第二个兼容缓存</t></text></comment>'
    + '<comment ref="D4" authorId="2"><text><t>独立经典备注</t></text></comment>'
    + '</commentList></comments>')));
  return writeOpcZip(files);
}
async function read(bytes: Uint8Array) { return projectXlsxRecords(await readXlsxRecords(bytes)).model; }
describe('unambiguous legacy thread placeholders', () => {
  it('maps one authoritative thread and retains genuine classic notes without duplicating compatibility text', async () => {
    const model = await read(source()), snapshot = model.snapshot();
    expect(snapshot.annotations).toHaveLength(3); expect(snapshot.metadata.workbook.comments?.map(comment => comment.format)).toEqual(['modern', 'modern', 'classic']);
    expect(snapshot.annotations[0]?.sourceId).toBe(root); expect(snapshot.annotations[0]?.replies).toHaveLength(1);
    expect(snapshot.content?.html).not.toContain('兼容缓存'); expect(snapshot.content?.html).not.toContain('只供旧应用');
    expect(model.getNode(snapshot.annotations[2]!.bodyNodeId!)?.text).toBe('独立经典备注');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('keeps unmatched, moved and copied compatibility cases explicit until recovery semantics are settled', async () => {
    for (const edit of [
      (xml: string) => xml.replace(root, reply),
      (xml: string) => xml.replace('ref="B2"', 'ref="B3"'),
      (xml: string) => xml.replace('</commentList>', '<comment ref="E5" authorId="0"><text><t>复制的缓存</t></text></comment></commentList>'),
      (xml: string) => xml.replace(/<comment ref="B2"[\s\S]*?<\/comment>/, ''),
      (xml: string) => xml.replace(/<commentList>[\s\S]*?<\/commentList>/, '<commentList/>'),
    ]) await expect(read(source(edit))).rejects.toMatchObject({ code: 'UnsupportedFeature', message: 'A legacy thread placeholder needs an explicit supported reconciliation policy.' });
  });
  it('still reads standalone native threads that have no legacy compatibility part', async () => {
    const model = await read(readFileSync('examples/xlsx/10-modern-threads.xlsx')); expect(model.snapshot().annotations).toHaveLength(2);
  });
});
