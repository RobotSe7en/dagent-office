import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@doffi/core';
import { importXlsx } from '@doffi/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const evidence = JSON.parse(readFileSync('docs/evidence/xlsx-picture-rotation.json', 'utf8')) as { cases: { source: { fill: Record<string, string> }; producer: { fill: Record<string, string> } }[] };
describe('native picture-fill rotation', () => {
  it.each(['', '-producer'])('retains%s actual fill flags independently from drawing rotation, flips and package resources', async suffix => {
    const bytes = readFileSync('packages/office/tests/fixtures/xlsx-picture-rotation' + suffix + '.xlsx');
    if (!suffix) expect(execFileSync('python3', ['-c', "import importlib.util,sys;s=importlib.util.spec_from_file_location('pictures','scripts/check-xlsx-picture-rotation.py');m=importlib.util.module_from_spec(s);s.loader.exec_module(m);sys.stdout.buffer.write(m.generate(m.inputs()))"])).toEqual(bytes);
    const model = openDocument(await importXlsx(bytes, { documentId: 'picture-rotation' }), { kind: 'workbook' }), snapshot = model.snapshot();
    expect(snapshot.metadata.workbook.objects).toHaveLength(20);
    for (const [index, sheet] of snapshot.metadata.workbook.sheets.entries()) {
      const object = snapshot.metadata.workbook.objects!.find(object => object.sheetId === sheet.sheetId)!;
      const native = evidence.cases[index]![suffix ? 'producer' : 'source'].fill.rotWithShape;
      expect(model.getNode(object.nodeId)!.attributes['data-office-image-rotate-with-shape']).toBe(native === undefined ? undefined : String(['1', 'true'].includes(native)));
    }
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(snapshot);
  });
  it('rejects malformed native fill Booleans instead of interpreting them as omitted defaults', async () => {
    const source = readDocumentZip(readFileSync('packages/office/tests/fixtures/xlsx-picture-rotation.xlsx')), path = 'xl/drawings/drawing3.xml';
    const original = new TextDecoder().decode(source.get(path)!);
    for (const invalid of ['yes', '2', '', 'NaN']) {
      const files = new Map(source); files.set(path, new TextEncoder().encode(original.replace('rotWithShape="0"', 'rotWithShape="' + invalid + '"')));
      await expect(importXlsx(writeOpcZip(files), { documentId: 'bad-picture' })).rejects.toMatchObject({ code: 'InvalidDocument' });
    }
  });
  it('validates preserved picture-fill flags again when opening the persisted HTML package', async () => {
    const files = new Map(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-picture-rotation.xlsx'), { documentId: 'picture-package' }));
    const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as { content: string };
    const html = new TextDecoder().decode(files.get(manifest.content)!);
    for (const invalid of ['0', '', 'yes']) {
      const changed = new Map(files); changed.set(manifest.content, new TextEncoder().encode(html.replace('data-office-image-rotate-with-shape="false"', 'data-office-image-rotate-with-shape="' + invalid + '"')));
      expect(() => openDocument(changed, { kind: 'workbook' })).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
    }
  });
});
