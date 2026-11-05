import { expect, it } from 'vitest';
import { createWorkbook, DocumentError, openDocument } from '@dagent-office/core';
import { AxisDimensions } from '../src/workbook/dimensions.js';
import { implicitWorkbookRowSources } from '../src/workbook/implicit-row-sources.js';
const source = openDocument(createWorkbook({ documentId: 'implicit-fonts' }), { kind: 'workbook' }).snapshot().metadata.workbook.sheets[0]!;
const columns = new AxisDimensions(16_384, 20, [{ from: 1, to: 1, sizePt: 0 }]);
it('finds native implicit typography around actual cells, hidden columns and merged spans, including empty styled rows', async () => {
  const sheet = { ...source, rowDimensions: [{ from: 1, to: 1, cssClass: 'row-font' }, { from: 10, to: 10, cssClass: 'row-font' }],
    columnDimensions: [{ from: 2, to: 4, unit: 'characters' as const, cssClass: 'column-font' }] };
  const occupied = new Map([[0, new Set([2, 3])], [1, new Set([0])], [2, new Set([2, 3, 4])]]), points = []; let admitted = 0;
  for await (const point of implicitWorkbookRowSources(sheet, columns, occupied, 2, true, () => true, (row, column) => row === 1 && column >= 2 && column <= 5 ? 5 : undefined,
    () => { admitted++; }, new AbortController().signal)) points.push(point);
  expect(points).toEqual([{ row: 0, column: 4 }, { row: 1, column: 6 }, { row: 10, column: 0 }]); expect(admitted).toBe(3);
  const unused = [];
  for await (const point of implicitWorkbookRowSources({ ...sheet, rowDimensions: [] }, columns, new Map(), -1, true, () => true, () => undefined, () => {}, new AbortController().signal)) unused.push(point);
  expect(unused).toEqual([]);
});
it('resolves every implicit coordinate for custom selectors and honors cumulative admission and cancellation', async () => {
  const sheet = { ...source, rowDimensions: [], columnDimensions: [{ from: 2, to: 4, unit: 'characters' as const, cssClass: 'column-font' }] }, points = [];
  for await (const point of implicitWorkbookRowSources(sheet, columns, new Map([[0, new Set([2])]]), 0, false, () => true, () => undefined, () => {}, new AbortController().signal)) points.push(point);
  expect(points).toEqual([{ row: 0, column: 3 }, { row: 0, column: 4 }]);
  let admitted = 0;
  const limited = implicitWorkbookRowSources(sheet, columns, new Map(), 1000, false, () => true, () => undefined, () => {
    if (++admitted > 1) throw new DocumentError('ResourceLimitExceeded', 'Test cumulative admission.');
  }, new AbortController().signal);
  expect((await limited.next()).value).toEqual({ row: 0, column: 2 }); await expect(limited.next()).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
  const abort = new AbortController(); abort.abort();
  const cancelled = implicitWorkbookRowSources(sheet, columns, new Map(), 1000, false, () => true, () => undefined, () => {}, abort.signal);
  await expect(cancelled.next()).rejects.toMatchObject({ code: 'OperationCancelled' });
});
