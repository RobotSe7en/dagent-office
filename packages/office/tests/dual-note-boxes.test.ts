import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import type { WorkbookComment } from '@dagent-office/core';
type ClassicComment = Extract<WorkbookComment, { format: 'classic' }>;
import { importXlsx } from '@dagent-office/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const cases = ['conflict', 'same-anchor', 'visible', 'commentPr-only'];
describe('classic dual-box source precedence', () => {
  it.each(cases)('keeps both %s native sources and selects VML on exact package reopen', async name => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-' + name + '.xlsx'), { documentId: 'dual-box-' + name }), { kind: 'workbook' });
    const snapshot = model.snapshot(), comment = snapshot.metadata.workbook.comments![0]!;
    expect(comment.format).toBe('classic'); if (comment.format !== 'classic') throw Error('A native classic note is required');
    expect(comment.boxSource).toBe(name === 'commentPr-only' ? 'commentPr' : 'vml');
    expect(model.getNode(snapshot.annotations[0]!.bodyNodeId!)!.text).toBe('请确认：预算😀禰\u{e0100}');
    if (name === 'commentPr-only') { expect(comment.box!.altText).toBe('备用备注框😀'); expect(comment.commentPrBox).toBeUndefined(); expect(comment.boxDifferences).toBeUndefined(); }
    else {
      expect(comment.box).toMatchObject({ visible: name === 'visible', printable: true, anchor: { kind: 'two-cell', from: { row: 1, column: 2, offsetXPt: 7.5, offsetYPt: 1.5 }, sizeWithCells: true } });
      expect(comment.box!.altText).toBeUndefined();
      expect(comment.commentPrBox).toMatchObject({ altText: '备用备注框😀', printable: false, visible: false });
      expect(comment.boxDifferences).toEqual(expect.arrayContaining(['printable', 'altText', 'appearance']));
      expect(comment.boxDifferences!.includes('anchor')).toBe(name !== 'same-anchor');
      expect(comment.boxDifferences!.includes('visible')).toBe(name === 'visible');
      expect(comment.box!.cssClass).not.toBe(comment.commentPrBox!.cssClass);
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it.each(cases)('reopens the independently observed %s producer without reconstructing discarded properties', async name => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-' + name + '-producer.xlsx'), { documentId: 'dual-producer-' + name }), { kind: 'workbook' });
    const comment = model.snapshot().metadata.workbook.comments![0]!;
    expect(comment.format).toBe('classic'); if (comment.format !== 'classic') throw Error('A classic note is required');
    expect(comment.boxSource).toBe('vml'); expect(comment.commentPrBox).toBeUndefined(); expect(comment.boxDifferences).toBeUndefined();
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
  });
  it('validates the unselected native box instead of dropping malformed or unsupported properties', async () => {
    for (const edit of [
      (s: string) => s.replace('textHAlign="center"', 'textHAlign="unsupported"'),
      (s: string) => s.replace('textHAlign="center"', 'unexpected="1"'),
      (s: string) => s.replace(/<(?:\w+:)?row>2<\/(?:\w+:)?row>/, '<xdr:row xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">1048576</xdr:row>'),
    ]) {
      const files = new Map(readDocumentZip(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx'))), path = 'xl/comments/comment1.xml';
      const source = new TextDecoder().decode(files.get(path)!); expect(edit(source)).not.toBe(source); files.set(path, new TextEncoder().encode(edit(source)));
      await expect(importXlsx(writeOpcZip(files), { documentId: 'invalid-unselected-box' })).rejects.toBeDefined();
    }
  });
  it('checks alternate style ownership, source consistency and unique difference reports while retaining old single-box packages', async () => {
    const files = await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-dual-note-boxes-conflict.xlsx'), { documentId: 'dual-semantic' });
    const original = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
    for (const edit of [
      (comment: ClassicComment) => { comment.commentPrBox!.cssClass = 'missing'; },
      (comment: ClassicComment) => { comment.boxSource = 'commentPr'; },
      (comment: ClassicComment) => { delete comment.box; },
      (comment: ClassicComment) => { comment.boxDifferences = ['anchor', 'anchor']; },
      (comment: ClassicComment) => { delete comment.commentPrBox; },
    ]) {
      const metadata = structuredClone(original); edit(metadata.workbook.comments[0]);
      const changed = new Map(files); changed.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
      expect(() => openDocument(changed, { kind: 'workbook' })).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
    }
    const metadata = structuredClone(original), comment = metadata.workbook.comments[0]; delete comment.boxSource; delete comment.commentPrBox; delete comment.boxDifferences;
    const legacy = new Map(files); legacy.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
    expect(openDocument(legacy, { kind: 'workbook' }).snapshot().metadata.workbook.comments![0]!.format).toBe('classic');
  });
});
