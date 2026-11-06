import { expect, it } from 'vitest';
import type { WorkbookCell } from '@doffi/core';
import { checkedPlainWorkbookAttributes, checkedPlainWorkbookCell } from '../src/workbook/plain-cell-source.js';
const text = '客户 123 😀禰󠄀';
const cell: WorkbookCell = { sheetId: 's', row: 1, column: 2, value: { type: 'text', value: text }, text, cssClass: 'normal',
  html: '<td data-column="2" data-value-type="text" class="normal">' + text + '</td>' };
it('validates canonical Unicode text, typed scalars and formula identities without changing bytes', () => {
  expect(checkedPlainWorkbookCell(cell)).toBe(true);
  const formula: WorkbookCell = { sheetId: cell.sheetId, row: cell.row, column: cell.column, text: '24', value: { type: 'number', value: 24 }, formula: { expression: '=SUM(A1:B1)', cacheState: 'imported' },
    html: '<td data-column="2" data-value-type="number" data-value="24" data-formula="=SUM(A1:B1)" data-cache-state="imported">24</td>' };
  expect(checkedPlainWorkbookCell(formula)).toBe(true); expect(cell.text).toBe(text);
});
it('keeps rich, entity-encoded and other valid HTML syntax on the inert parser path', () => {
  for (const html of [cell.html.replace(text, '<span>' + text + '</span>'), cell.html.replace('123', '&#49;23'), cell.html.replaceAll('"', "'"), cell.html.replace('</td>', '</TD>'), cell.html.replace('123', '1\r\n23'), cell.html.replace('<td ', '<td\u00a0'), cell.html.replace(' class=', '\u00a0class=')])
    expect(checkedPlainWorkbookCell({ ...cell, html })).toBe(false);
});
it('retains lexical attributes for equivalent DOM construction and CSS selectors', () => {
  const number: WorkbookCell = { sheetId: 's', row: 1, column: 2, value: { type: 'number', value: 1000 }, text: '1000',
    html: '<td\tdata-column="2" data-value-type="number" data-value="1e3" >1000</td>' };
  expect([...checkedPlainWorkbookAttributes(number)!]).toEqual([['data-column', '2'], ['data-value-type', 'number'], ['data-value', '1e3']]);
});
it('rejects handlers, resources, duplicate fields and inconsistent source facts even for a previously seen format', () => {
  for (const html of [cell.html.replace('<td ', '<td onclick="run()" '), cell.html.replace('<td ', '<td src="https://example.com/private" '),
    cell.html.replace('<td ', '<td data-column="2" '), cell.html.replace('data-column="2"', 'data-column="3"'), cell.html.replace(text, 'different'), cell.html.replace('class="normal"', 'class="other"')])
    expect(() => checkedPlainWorkbookCell({ ...cell, html })).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
  expect(() => checkedPlainWorkbookCell({ ...cell, value: { type: 'text', value: 'different' } })).toThrowError(expect.objectContaining({ code: 'InvalidDocument' }));
});
