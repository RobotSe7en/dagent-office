import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { openDocument } from '@dagent-office/core';
import { importPptx, importXlsx } from '@dagent-office/office';
import { readDocumentZip, writeOpcZip } from '../src/document-zip.js';

const decoder = new TextDecoder(), encoder = new TextEncoder();
function replace(bytes: Uint8Array, path: string, edit: (xml: string) => string) {
  const files = new Map(readDocumentZip(bytes));
  files.set(path, encoder.encode(edit(decoder.decode(files.get(path)!)))); return writeOpcZip(files);
}
const chartNamespace = 'http://schemas.openxmlformats.org/drawingml/2006/chart';

describe('independent PR25 native-reading review regressions', () => {
  it.each([false, true])('preserves adjacent fixed-default/automatic rows in either order (%s)', async reverse => {
    const flags = reverse ? [false, true, false, false] : [true, false, true, true];
    const bytes = replace(readFileSync('examples/xlsx/13-date-1904.xlsx'), 'xl/worksheets/sheet1.xml', () =>
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetFormatPr defaultRowHeight="30" customHeight="1"/><sheetData>'
      + flags.map((auto, index) => `<row r="${index + 1}" hidden="0"${auto ? ' customHeight="0"' : ''}><c r="A${index + 1}" t="inlineStr"><is><t>Row ${index + 1}</t></is></c></row>`).join('')
      + '</sheetData></worksheet>');
    const model = openDocument(await importXlsx(bytes, { documentId: 'review-rows' }), { kind: 'workbook' });
    const sheet = model.snapshot().metadata.workbook.sheets[0]!;
    expect(sheet.defaultRowHeightCustom).toBe(true);
    expect(flags.map((_, row) => sheet.rowDimensions.find(run => run.from <= row && run.to >= row)?.autoHeight ?? false)).toEqual(flags);
    expect(sheet.rowDimensions.map(run => [run.from, run.to])).toEqual(reverse ? [[1, 1]] : [[0, 0], [2, 3]]);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
  });
  it.each([['xlsx', 'prefix'], ['xlsx', 'default'], ['pptx', 'prefix'], ['pptx', 'default']])('accepts local chart %s/%s namespace declarations', async (kind, binding) => {
    const path = kind === 'xlsx' ? 'xl/charts/chart2.xml' : 'ppt/charts/chart2.xml';
    const original = readFileSync(kind === 'xlsx' ? 'examples/xlsx/07-three-charts.xlsx' : 'examples/pptx/09-charts.pptx');
    const bytes = replace(original, path, xml => binding === 'prefix'
      ? xml.replace(/<(c:(?:marker|symbol|size|crosses|crossesAt))(?=[\s/>])/g, `<$1 xmlns:c="${chartNamespace}"`)
      : xml.replace(/<c:(marker|symbol|size|crosses|crossesAt)(?=[\s/>])/g, `<$1 xmlns="${chartNamespace}"`).replace(/<\/c:(marker|symbol|size|crosses|crossesAt)>/g, '</$1>'));
    if (kind === 'xlsx') {
      const expected = openDocument(await importXlsx(original, { documentId: 'local-ns' }), { kind: 'workbook' });
      const actual = openDocument(await importXlsx(bytes, { documentId: 'local-ns' }), { kind: 'workbook' });
      expect(actual.snapshot()).toEqual(expected.snapshot());
    } else expect(openDocument(importPptx(bytes, { documentId: 'local-ns' }), { kind: 'presentation' }).snapshot())
      .toEqual(openDocument(importPptx(original, { documentId: 'local-ns' }), { kind: 'presentation' }).snapshot());
  });
  it('preserves the native default zero span of a manual page break', async () => {
    const bytes = replace(readFileSync('examples/xlsx/11-print-layout.xlsx'), 'xl/worksheets/sheet1.xml', xml =>
      xml.replace(/<rowBreaks\b[\s\S]*?<\/rowBreaks>/, '<rowBreaks count="1" manualBreakCount="1"><brk id="23" man="1"/></rowBreaks>'));
    const model = openDocument(await importXlsx(bytes, { documentId: 'review-break' }), { kind: 'workbook' });
    expect(model.snapshot().metadata.workbook.sheets[0]!.print!.rowBreaks).toEqual([{ before: 23, from: 0, to: 0 }]);
    expect(openDocument(model.serialize(), { kind: 'workbook' }).snapshot()).toEqual(model.snapshot());
  });
});
