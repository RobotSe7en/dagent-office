import { createHash } from "node:crypto";
import { describe, expect, it } from "vitest";
import { createWorkbook, openDocument } from "@dagent-office/core";
import type { WorkbookDocumentMetadata, WorkbookManifest } from "@dagent-office/core";
import { AxisDimensions } from "../src/workbook/dimensions.js";
import { modelWindowSource } from "../src/workbook/window-source.js";
import type { WorkbookWindowRequest, WorkbookWindowResponse, WorkbookWindowSource } from "../src/workbook/window-source.js";
import { WorkbookWindowLoader } from "../src/workbook/window-loader.js";

const encode = (value: string) => new TextEncoder().encode(value);
const dimensions = { rows: new AxisDimensions(1_048_576, 10, []), columns: new AxisDimensions(16_384, 20, []) };
const viewport = { widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 };
function source(merged = false) {
  const files = new Map(createWorkbook({ documentId: 'windows' }));
  const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as WorkbookManifest;
  const metadata = JSON.parse(new TextDecoder().decode(files.get(manifest.metadata)!)) as WorkbookDocumentMetadata;
  if (merged) {
    metadata.workbook.sheets[0]!.merges = [{ from: { row: 0, column: 0 }, to: { row: 6, column: 6 } }];
    metadata.workbook.sheets[0]!.frozen = { rows: 2, columns: 1 };
  }
  const path = 'sheets/sheet-1/0-0.html';
  const html = '<!doctype html><html><head></head><body><table data-sheet-id="sheet-1" data-row="0" data-column="0"><tbody><tr data-row="0"><td data-column="0" data-value-type="text">首格😀</td></tr>'
    + (merged ? '<tr data-row="6"><td data-column="6" data-value-type="blank" class="cell-default"></td></tr>' : '')
    + '<tr data-row="9"><td data-column="9" data-value-type="number" data-value="42" data-formula="=40+2" data-cache-state="imported">42</td></tr></tbody></table></body></html>';
  const bytes = encode(html); files.set(path, bytes);
  manifest.chunks.push({ sheetId: 'sheet-1', row: 0, column: 0, path, sha256: createHash('sha256').update(bytes).digest('hex') });
  files.set('manifest.json', encode(JSON.stringify(manifest))); files.set(manifest.metadata, encode(JSON.stringify(metadata)));
  return modelWindowSource(openDocument(files, { kind: 'workbook' }));
}
function queued(input: WorkbookWindowSource) {
  const requests: { request: WorkbookWindowRequest; signal: AbortSignal; resolve(value: WorkbookWindowResponse): void; reject(error: Error): void }[] = [];
  const provider: WorkbookWindowSource = { snapshot: input.snapshot, read: (request, signal) => new Promise((resolve, reject) => requests.push({ request, signal, resolve, reject })) };
  const complete = async (index: number) => { const entry = requests[index]!; entry.resolve(await input.read(entry.request, new AbortController().signal)); };
  return { provider, requests, complete };
}
describe('bounded immutable workbook viewport reads', () => {
  it('discovers offscreen anchors through actual row chunks in separately bounded endpoint batches', async () => {
    const input = source(), requests: WorkbookWindowRequest[] = [], loader = new WorkbookWindowLoader({ snapshot: input.snapshot,
      read(request, signal) { requests.push(request); return input.read(request, signal); } }, { maxSourceCells: 100, maxRowReadCells: 1000 }, true);
    expect(await loader.load('sheet-1', dimensions, { ...viewport, scrollLeftPt: 10_000, scrollTopPt: 0 })).toBe(true);
    const frame = loader.frame!;
    expect(frame.projections.every(cell => cell.column > 400)).toBe(true); expect(frame.getCell(0, 0)?.text).toBe('首格😀');
    expect(frame.completeRows).toEqual([0, 1, 2, 3, 4, 5, 6, 7]); expect(frame.storedCellCount).toBe(1);
    expect(requests.length).toBeGreaterThan(1);
    expect(requests.every(request => request.windows.reduce((count, window) => count + window.rows * window.columns, 0) <= 100)).toBe(true);
    expect(requests.reduce((sum, request) => sum + request.windows.reduce((count, window) => count + window.rows * window.columns, 0), 0)).toBeGreaterThan(frame.sourceCoordinates.length);
    expect(frame.sourceCoordinates.length).toBeLessThanOrEqual(100); expect(Object.isFrozen(frame.completeRows)).toBe(true);
  });
  it('checks row-discovery slots before reading and rejects responses that belong to another batch', async () => {
    const input = source(), requests: WorkbookWindowRequest[] = [], bounded = new WorkbookWindowLoader({ snapshot: input.snapshot,
      read(request, signal) { requests.push(request); return input.read(request, signal); } }, { maxSourceCells: 100, maxRowReadCells: 10 }, true);
    const small = { widthPt: 20, heightPt: 10, scrollLeftPt: 10_000, scrollTopPt: 6000 };
    expect(await bounded.load('sheet-1', dimensions, small)).toBe(true); const accepted = bounded.frame, count = requests.length;
    await expect(bounded.load('sheet-1', dimensions, { ...small, scrollTopPt: 0 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(requests).toHaveLength(count); expect(bounded.frame).toBe(accepted);
    const later = (await input.read({ documentId: 'windows', revision: 0, windows: [{ sheetId: 'sheet-1', row: 9, column: 9, rows: 1, columns: 1 }], maxCells: 1, maxCodeUnits: 8192 }, new AbortController().signal)).cells[0]!;
    const mixed = new WorkbookWindowLoader({ snapshot: input.snapshot, async read(request, signal) { const response = await input.read(request, signal); return { ...response, cells: [...response.cells, later] }; } }, { maxSourceCells: 100 }, true);
    await expect(mixed.load('sheet-1', dimensions, { ...small, widthPt: 100, heightPt: 100, scrollTopPt: 0 })).rejects.toMatchObject({ code: 'InvalidDocument' });
    expect(mixed.frame).toBeUndefined();
  });
  it('keeps one content anchor for four merged pane projections and also reads covered border sources', async () => {
    const input = source(true), queries: WorkbookWindowRequest[] = [];
    const loader = new WorkbookWindowLoader({ snapshot: input.snapshot, read(request, signal) { queries.push(request); return input.read(request, signal); } }, { maxCells: 40 });
    expect(await loader.load('sheet-1', dimensions, viewport)).toBe(true);
    const frame = loader.frame!;
    expect(frame.projections.filter(cell => cell.row === 0 && cell.column === 0)).toHaveLength(4);
    expect(queries[0]!.windows.reduce((sum, window) => sum + window.rows * window.columns, 0)).toBe(frame.sourceCoordinates.length);
    expect(frame.sourceCoordinates.length).toBeGreaterThan(21); expect(frame.sourceCoordinates.length).toBeLessThan(100);
    expect(frame.storedCellCount).toBe(2); expect(frame.getCell(0, 0)?.text).toBe('首格😀');
    expect(frame.getCell(6, 6)).toMatchObject({ value: { type: 'blank' }, cssClass: 'cell-default' });
    expect(frame.getCell(7, 3)).toBeUndefined(); expect(frame.getCell(-1, 16_384)).toBeUndefined();
    expect(Object.isFrozen(frame.projections[0]?.clip)).toBe(true); expect(Object.isFrozen(frame.layout.panes[0]?.rows)).toBe(true);
    expect(Object.isFrozen(frame.getCell(0, 0)?.value)).toBe(true); expect(Object.isFrozen(queries[0]?.windows)).toBe(true);
  });
  it('rejects late windows even if the provider ignores cancellation and preserves formula cache identity', async () => {
    const queue = queued(source()), loader = new WorkbookWindowLoader(queue.provider);
    const first = loader.load('sheet-1', dimensions, { ...viewport, scrollLeftPt: 0, scrollTopPt: 0 });
    const second = loader.load('sheet-1', dimensions, { ...viewport, scrollLeftPt: 180, scrollTopPt: 90 });
    expect(queue.requests[0]!.signal.aborted).toBe(true);
    await queue.complete(1); expect(await second).toBe(true); const accepted = loader.frame;
    expect(accepted?.getCell(9, 9)?.formula).toEqual({ expression: '=40+2', cacheState: 'imported' });
    await queue.complete(0); expect(await first).toBe(false); expect(loader.frame).toBe(accepted);
  });
  it('ignores obsolete failures and disposal responses, but propagates a current read failure atomically', async () => {
    const queue = queued(source()), loader = new WorkbookWindowLoader(queue.provider);
    const old = loader.load('sheet-1', dimensions, viewport), current = loader.load('sheet-1', dimensions, viewport);
    queue.requests[0]!.reject(new Error('obsolete')); expect(await old).toBe(false);
    await queue.complete(1); expect(await current).toBe(true); const frame = loader.frame;
    const failing = loader.load('sheet-1', dimensions, viewport);
    queue.requests[2]!.reject(new Error('current failure')); await expect(failing).rejects.toThrow('current failure'); expect(loader.frame).toBe(frame);
    const late = loader.load('sheet-1', dimensions, viewport); loader.dispose();
    expect(queue.requests[3]!.signal.aborted).toBe(true); await queue.complete(3); expect(await late).toBe(false); expect(loader.frame).toBeUndefined();
    await expect(loader.load('sheet-1', dimensions, viewport)).rejects.toMatchObject({ code: 'InvalidOperation' });
  });
  it('rejects mismatched revision, duplicate and unrequested response coordinates without publishing a partial frame', async () => {
    for (const corrupt of [
      (response: WorkbookWindowResponse) => ({ ...response, revision: response.revision + 1 }),
      (response: WorkbookWindowResponse) => ({ ...response, documentId: 'elsewhere' }),
      (response: WorkbookWindowResponse) => ({ ...response, cells: [...response.cells, ...response.cells] }),
      (response: WorkbookWindowResponse) => ({ ...response, cells: response.cells.map(cell => ({ ...cell, row: 100 })) }),
    ]) {
      const input = source(), loader = new WorkbookWindowLoader({ snapshot: input.snapshot, read: async (request, signal) => corrupt(await input.read(request, signal)) });
      await expect(loader.load('sheet-1', dimensions, { ...viewport, scrollLeftPt: 0, scrollTopPt: 0 })).rejects.toBeDefined();
      expect(loader.frame).toBeUndefined();
    }
  });
  it('checks cumulative windows before a model read and enforces text limits on both provider and receiver', async () => {
    const input = source(), identity = { documentId: 'windows', revision: input.snapshot.manifest.revision }, signal = new AbortController().signal;
    const window = { sheetId: 'sheet-1', row: 0, column: 0, rows: 1, columns: 1 };
    await expect(input.read({ ...identity, windows: [window, { ...window, row: 1 }], maxCells: 1, maxCodeUnits: 1000 }, signal)).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    await expect(input.read({ ...identity, windows: [window, window], maxCells: 2, maxCodeUnits: 1000 }, signal)).rejects.toMatchObject({ code: 'InvalidOperation' });
    await expect(input.read({ ...identity, windows: [window], maxCells: 1, maxCodeUnits: 1 }, signal)).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    const controller = new AbortController(); controller.abort();
    await expect(input.read({ ...identity, windows: [window], maxCells: 1, maxCodeUnits: 1000 }, controller.signal)).rejects.toMatchObject({ code: 'OperationCancelled' });
    const response = await input.read({ ...identity, windows: [window], maxCells: 1, maxCodeUnits: 1000 }, signal);
    const loader = new WorkbookWindowLoader({ snapshot: input.snapshot, read: async () => response }, { maxCodeUnits: 1 });
    await expect(loader.load('sheet-1', dimensions, { ...viewport, scrollLeftPt: 0, scrollTopPt: 0 })).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(loader.frame).toBeUndefined();
  });
  it('does not request any data for a zero-sized viewport', async () => {
    const queue = queued(source()), loader = new WorkbookWindowLoader(queue.provider);
    expect(await loader.load('sheet-1', dimensions, { ...viewport, widthPt: 0 })).toBe(true);
    expect(queue.requests).toHaveLength(0); expect(loader.frame?.projections).toEqual([]);
  });
  it('checks the separate appearance-source budget before sending a query', async () => {
    const queue = queued(source(true)), loader = new WorkbookWindowLoader(queue.provider, { maxCells: 40, maxSourceCells: 21 });
    await expect(loader.load('sheet-1', dimensions, viewport)).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(queue.requests).toHaveLength(0); expect(loader.frame).toBeUndefined();
    expect(() => new WorkbookWindowLoader(queue.provider, { maxSourceCells: 0 })).toThrowError(expect.objectContaining({ code: 'InvalidOperation' }));
  });
  it('retains the accepted frame when a later viewport exceeds its appearance budget', async () => {
    const input = source(), requests: WorkbookWindowRequest[] = [], loader = new WorkbookWindowLoader({ snapshot: input.snapshot, read(request, signal) { requests.push(request); return input.read(request, signal); } }, { maxSourceCells: 5 });
    expect(await loader.load('sheet-1', dimensions, { widthPt: 20, heightPt: 10, scrollLeftPt: 0, scrollTopPt: 0 })).toBe(true);
    const accepted = loader.frame;
    await expect(loader.load('sheet-1', dimensions, viewport)).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
    expect(requests).toHaveLength(1); expect(loader.frame).toBe(accepted);
  });
  it('yields bounded subwindows so cancellation interrupts a direct large read before its next allocation', async () => {
    const model = openDocument(createWorkbook({ documentId: 'cancel' }), { kind: 'workbook' });
    const queried: number[] = [], input = modelWindowSource({ ...model, readWindow(window) { queried.push(window.rows * window.columns); return model.readWindow(window); } });
    const controller = new AbortController(), pending = input.read({ documentId: 'cancel', revision: model.snapshot().manifest.revision,
      windows: [{ sheetId: 'sheet-1', row: 0, column: 0, rows: 1, columns: 1000 }], maxCells: 1000, maxCodeUnits: 1000 }, controller.signal);
    expect(queried).toEqual([256]); controller.abort();
    await expect(pending).rejects.toMatchObject({ code: 'OperationCancelled' }); expect(queried).toEqual([256]);
  });
  it('isolates the requested windows before yielding to caller code', async () => {
    const model = openDocument(createWorkbook({ documentId: 'isolated' }), { kind: 'workbook' });
    const queried: number[] = [], input = modelWindowSource({ ...model, readWindow(window) { queried.push(window.column); return model.readWindow(window); } });
    const windows = [{ sheetId: 'sheet-1', row: 0, column: 0, rows: 1, columns: 512 }];
    const pending = input.read({ documentId: 'isolated', revision: model.snapshot().manifest.revision, windows, maxCells: 512, maxCodeUnits: 1000 }, new AbortController().signal);
    windows[0]!.columns = 16_384; windows[0]!.column = 1000;
    await pending; expect(queried).toEqual([0, 256]);
  });
});
