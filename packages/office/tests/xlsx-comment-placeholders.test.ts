import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { openDocument } from "@doffi/core";
import { readDocumentZip, writeOpcZip } from "../src/document-zip.js";
import { readXlsxRecords } from "../src/xlsx/read.js";
import { projectXlsxRecords } from "./helpers/xlsx-record-model.js";

const encoder = new TextEncoder(), decoder = new TextDecoder();
const root = '{33333333-3333-4333-8333-333333333333}', reply = '{44444444-4444-4444-8444-444444444444}', second = '{55555555-5555-4555-8555-555555555555}';
function source(edit: (xml: string) => string = value => value, editNative: (files: Map<string, Uint8Array>) => void = () => {}): Uint8Array {
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
  editNative(files); return writeOpcZip(files);
}
async function read(bytes: Uint8Array) { return projectXlsxRecords(await readXlsxRecords(bytes)).model; }
describe('legacy thread placeholder reconciliation', () => {
  it('preserves revision UIDs separately from classic GUIDs and modern source identities, independent of the namespace prefix', async () => {
    const uid = '{DDDDDDDD-DDDD-4DDD-8DDD-DDDDDDDDDDDD}', guid = '{EEEEEEEE-EEEE-4EEE-8EEE-EEEEEEEEEEEE}';
    const model = await read(source(xml => xml.replace('<comments ', '<comments xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:revision="http://schemas.microsoft.com/office/spreadsheetml/2014/revision" mc:Ignorable="revision" ')
      .replace('<comment ref="B2"', `<comment revision:uid="${root}" ref="B2"`).replace('<comment ref="C3"', `<comment revision:uid="${uid}" ref="C3"`)
      .replace('<comment ref="D4"', `<comment revision:uid="${uid}" guid="${guid}" ref="D4"`)));
    const snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.comments?.[0]).toMatchObject({ compatibilityUids: [root] }); expect(snapshot.annotations[0]!.status).toBe('resolved');
    expect(snapshot.metadata.workbook.comments?.[1]).toMatchObject({ compatibilityUids: [uid], recoveryReasons: ['mismatched-placeholder-identity'], recoveryStatus: 'open' });
    expect(snapshot.annotations[1]!.status).toBe('needs_reanchor');
    expect(snapshot.metadata.workbook.comments?.[2]).toMatchObject({ nativeUid: uid }); expect(snapshot.annotations[2]!.sourceId).toBe(guid);
    expect(openDocument(model.serialize()).snapshot()).toEqual(snapshot);
  });
  it.each(['bad-uid', 'unknown-attribute', 'unknown-namespace', 'unknown-mc-attribute', 'non-xml-whitespace'])('rejects %s despite an Ignorable declaration', async kind => {
    const bytes = source(xml => {
      const uid = kind === 'bad-uid' ? 'bad' : root, namespace = kind === 'unknown-namespace' ? 'urn:unselected-revision' : 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision';
      return xml.replace('<comments ', `<comments xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:rev="${namespace}" mc:Ignorable="${kind === 'non-xml-whitespace' ? '\u00a0' : ''}rev"${kind === 'unknown-mc-attribute' ? ' mc:PreserveAttributes="rev:uid"' : ''} `)
        .replace('<comment ref="B2"', `<comment rev:uid="${uid}"${kind === 'unknown-attribute' ? ' rev:unknown="1"' : ''} ref="B2"`);
    });
    await expect(read(bytes)).rejects.toMatchObject({ code: kind === 'bad-uid' ? 'InvalidDocument' : 'UnsupportedFeature' });
  });
  it('maps one authoritative thread and retains genuine classic notes without duplicating compatibility text', async () => {
    const model = await read(source()), snapshot = model.snapshot();
    expect(snapshot.annotations).toHaveLength(3); expect(snapshot.metadata.workbook.comments?.map(comment => comment.format)).toEqual(['modern', 'modern', 'classic']);
    expect(snapshot.annotations[0]?.sourceId).toBe(root); expect(snapshot.annotations[0]?.replies).toHaveLength(1);
    expect(snapshot.content?.html).not.toContain('兼容缓存'); expect(snapshot.content?.html).not.toContain('只供旧应用');
    expect(model.getNode(snapshot.annotations[2]!.bodyNodeId!)?.text).toBe('独立经典备注');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it.each([
    { name: 'moved', edit: (xml: string) => xml.replace('ref="B2"', 'ref="B3"'), reasons: ['mismatched-placeholder'], count: 3 },
    { name: 'copied', edit: (xml: string) => xml.replace('</commentList>', '<comment ref="E5" authorId="0"><text><t>复制的缓存</t></text></comment></commentList>'), reasons: ['duplicate-placeholder', 'mismatched-placeholder'], count: 3 },
    { name: 'deleted', edit: (xml: string) => xml.replace(/<comment ref="B2"[\s\S]*?<\/comment>/, ''), reasons: ['missing-placeholder'], count: 3 },
    { name: 'empty compatibility part', edit: (xml: string) => xml.replace(/<commentList>[\s\S]*?<\/commentList>/, '<commentList/>'), reasons: ['missing-placeholder'], count: 2 },
  ])('preserves the authoritative thread and original status for a $name placeholder', async ({ edit, reasons, count }) => {
    const model = await read(source(edit)), snapshot = model.snapshot(), annotation = snapshot.annotations[0]!;
    expect(snapshot.annotations).toHaveLength(count); expect(snapshot.metadata.workbook.comments?.filter(c => c.format === 'modern')).toHaveLength(2);
    expect(annotation).toMatchObject({ sourceId: root, status: 'needs_reanchor', anchor: { kind: 'cells', sheetId: 'sheet-1', ranges: [{ from: { row: 1, column: 1 }, to: { row: 1, column: 1 } }] } });
    expect(snapshot.metadata.workbook.comments?.[0]).toMatchObject({ recoveryStatus: 'resolved', recoveryReasons: reasons });
    expect(model.getNode(annotation.bodyNodeId!)?.text).toBe('请复核预算😀禰󠄀');
    expect(model.getNode(annotation.replies![0]!.bodyNodeId!)?.text).toBe('已复核并补充。');
    expect(snapshot.content?.html).not.toContain('兼容缓存'); expect(snapshot.content?.html).not.toContain('复制的缓存');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('preserves orphaned compatibility text as an invalidated classic note instead of inventing a modern thread', async () => {
    const model = await read(source(xml => xml.replace(root, reply))), snapshot = model.snapshot();
    expect(snapshot.annotations).toHaveLength(4); expect(snapshot.metadata.workbook.comments?.map(c => c.format)).toEqual(['modern', 'modern', 'classic', 'classic']);
    expect(snapshot.metadata.workbook.comments?.[0]).toMatchObject({ recoveryStatus: 'resolved', recoveryReasons: ['missing-placeholder'] });
    expect(snapshot.metadata.workbook.comments?.[2]).toMatchObject({ compatibilityThreadId: reply, recoveryStatus: 'open', recoveryReasons: ['orphan-placeholder'] });
    expect(snapshot.annotations[2]?.status).toBe('needs_reanchor'); expect(model.getNode(snapshot.annotations[2]!.bodyNodeId!)?.text).toBe('只供旧应用使用的缓存');
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('keeps a missing modern location empty even if a compatibility note supplies a candidate address', async () => {
    const model = await read(source(undefined, files => {
      const path = 'xl/threadedComments/threadedComment1.xml'; files.set(path, encoder.encode(decoder.decode(files.get(path)!).replace('ref="B2" ', '')));
    })), snapshot = model.snapshot();
    expect(snapshot.annotations[0]).toMatchObject({ status: 'needs_reanchor', anchor: { kind: 'cells', sheetId: 'sheet-1', ranges: [] } });
    expect(snapshot.metadata.workbook.comments?.[0]).toMatchObject({ recoveryStatus: 'resolved', recoveryReasons: ['missing-location'] });
    expect(snapshot.annotations[0]?.replies).toHaveLength(1);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('validates recovery provenance and status again on package reopen', async () => {
    const model = await read(source(xml => xml.replace(root, reply)));
    for (const kind of ['status', 'duplicate-reason', 'wrong-kind', 'missing-provenance', 'invented-cell']) {
      const files = new Map(model.serialize()), metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)), annotations = JSON.parse(decoder.decode(files.get('annotations.json')!));
      if (kind === 'status') annotations[0].status = 'resolved';
      else if (kind === 'duplicate-reason') metadata.workbook.comments[0].recoveryReasons.push('missing-placeholder');
      else if (kind === 'wrong-kind') metadata.workbook.comments[0].recoveryReasons = ['orphan-placeholder'];
      else if (kind === 'missing-provenance') delete metadata.workbook.comments[2].compatibilityThreadId;
      else metadata.workbook.comments[0].recoveryReasons = ['missing-location'];
      files.set('metadata.json', encoder.encode(JSON.stringify(metadata))); files.set('annotations.json', encoder.encode(JSON.stringify(annotations)));
      expect(() => openDocument(files, { kind: 'workbook' })).toThrow();
    }
  });
  it('still reads standalone native threads that have no legacy compatibility part', async () => {
    const model = await read(readFileSync('examples/xlsx/10-modern-threads.xlsx')); expect(model.snapshot().annotations).toHaveLength(2);
  });
});
