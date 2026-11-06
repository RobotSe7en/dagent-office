import { readFileSync } from 'node:fs';
import { describe, expect, it } from 'vitest';
import { createWorkbook, DEFAULT_WORKBOOK_RULE_LIMITS, openDocument, readWorkbookRules } from '@dagent-office/core';
import type { WorkbookModel } from '@dagent-office/core';
import { importXlsx } from '../../office/src/xlsx-import.js';
import { connectWorkbookReads } from '../src/workbook/read-client.js';
import { serveWorkbookReads } from '../src/workbook/read-server.js';
import { DEFAULT_CHART_READ_LIMITS, modelChartSource } from '../src/workbook/chart-source.js';
import { WORKBOOK_READ_PROTOCOL } from '../src/workbook/read-transport.js';

const empty = () => openDocument(createWorkbook({ documentId: 'transport' }), { kind: 'workbook' });
const tick = () => new Promise<void>(resolve => setTimeout(resolve, 5));
const request = (model: WorkbookModel, maxCells = 20) => ({ documentId: model.snapshot().manifest.documentId, revision: model.snapshot().manifest.revision,
  windows: [{ sheetId: model.snapshot().metadata.workbook.activeSheetId, row: 0, column: 0, rows: 1, columns: 20 }], maxCells, maxCodeUnits: 8192 });
const signal = () => new AbortController().signal;

describe('host-owned immutable workbook read transport', () => {
  it('retains native legacy marker flags in the frozen authorized snapshot without transferring source chunks', async () => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-auto-markers-producer.xlsx'), { documentId: 'marker-flags' }), { kind: 'workbook' });
    const ports = new MessageChannel(), server = serveWorkbookReads(ports.port1, { ...model, serialize() { throw Error('No full source transfer'); } }), client = await connectWorkbookReads(ports.port2);
    try {
      const charts = client.cells.snapshot.metadata.workbook.charts!;
      expect(charts).toHaveLength(8); expect(charts.every(chart => chart.nativeShowMarkers === true && Object.isFrozen(chart))).toBe(true);
      expect(charts[0]!.seriesStyles!.every(series => series.marker === 'none')).toBe(true);
      expect(client.cells.snapshot).toEqual(model.snapshot());
    } finally { client.dispose(); server.dispose(); }
  });
  it('transports current chart formulas and enforces both formula budgets before querying the model', async () => {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-formulas.xlsx'), { documentId: 'chart-transport' }), { kind: 'workbook' });
    let reads = 0;
    const ports = new MessageChannel(), server = serveWorkbookReads(ports.port1, { ...model, readWindow(window) { reads++; return model.readWindow(window); } },
      { readLimits: { maxChartFormulaCells: 1000, maxChartFormulaSteps: 10_000 } });
    const client = await connectWorkbookReads(ports.port2);
    const input = { ...DEFAULT_CHART_READ_LIMITS, documentId: 'chart-transport', revision: 0, chartIds: model.snapshot().metadata.workbook.charts!.map(chart => chart.chartId) };
    try {
      await expect(client.charts.read(input, signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' }); expect(reads).toBe(0);
      await expect(client.charts.read({ ...input, maxFormulaCells: 1000 }, signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' }); expect(reads).toBe(0);
      const accepted = { ...input, maxFormulaCells: 1000, maxFormulaSteps: 10_000 };
      expect(await client.charts.read(accepted, signal())).toEqual(await modelChartSource(model).read(accepted, signal()));
      expect(reads).toBeGreaterThan(0);
    } finally { client.dispose(); server.dispose(); }
  });
  it('transports current rule states through the same authorized revision and enforces cumulative server budgets', async () => {
    const model = openDocument(await importXlsx(readFileSync('examples/xlsx/06-rules-validation.xlsx'), { documentId: 'transport' }), { kind: 'workbook' }), before = model.serialize();
    const ports = new MessageChannel(), server = serveWorkbookReads(ports.port1, model, { readLimits: { maxRuleValues: 50 } }), client = await connectWorkbookReads(ports.port2);
    const input = { documentId: 'transport', revision: 0, sheetId: 'sheet-1', coordinates: [{ row: 2, column: 2 }, { row: 2, column: 4 }], ...DEFAULT_WORKBOOK_RULE_LIMITS };
    try {
      await expect(client.rules!.read(input, signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
      const admitted = { ...input, maxValues: 50 };
      expect(await client.rules!.read(admitted, signal())).toEqual(await readWorkbookRules(model, admitted));
      await expect(client.rules!.read({ ...admitted, revision: 1 }, signal())).rejects.toMatchObject({ code: 'RevisionConflict' });
      await expect(client.rules!.read({ ...admitted, coordinates: [{ row: -1, column: 0 }] }, signal())).rejects.toMatchObject({ code: 'InvalidOperation' });
      expect(model.serialize()).toEqual(before);
    } finally { client.dispose(); server.dispose(); }
  });
  it('transports native windows, current chart values and copied images without serializing the package', async () => {
    for (const name of ['07-three-charts', '08-images-shapes']) {
      const model = openDocument(await importXlsx(readFileSync(`examples/xlsx/${name}.xlsx`), { documentId: 'transport' }), { kind: 'workbook' });
      const ports = new MessageChannel(), server = serveWorkbookReads(ports.port1, { ...model, serialize() { throw Error('No full source transfer'); } });
      const client = await connectWorkbookReads(ports.port2);
      try {
        expect(client.cells.snapshot).toEqual(model.snapshot()); expect(client.cells.snapshot).not.toBe(model.snapshot());
        expect(Object.isFrozen(client.cells.snapshot.metadata.workbook.sheets[0])).toBe(true);
        const window = request(model);
        expect((await client.cells.read(window, signal())).cells).toEqual(model.readWindow(window.windows[0]!));
        const charts = { ...DEFAULT_CHART_READ_LIMITS, documentId: window.documentId, revision: window.revision, chartIds: model.snapshot().metadata.workbook.charts?.map(chart => chart.chartId) ?? [] };
        expect(await client.charts.read(charts, signal())).toEqual(await modelChartSource(model).read(charts, signal()));
        for (const asset of model.snapshot().manifest.assets) {
          const value = await client.assets.readAsset({ documentId: window.documentId, revision: window.revision, path: asset.path, maxBytes: 1024 * 1024 }, signal());
          expect(value.bytes).toEqual(model.getAsset(asset.path)); value.bytes.fill(0); expect(model.getAsset(asset.path)!.some(byte => byte !== 0)).toBe(true);
        }
      } finally { client.dispose(); server.dispose(); }
    }
  });

  it('enforces identities and both endpoint budgets before querying model cells', async () => {
    const model = empty(), ports = new MessageChannel(); let reads = 0;
    const server = serveWorkbookReads(ports.port1, { ...model, readWindow(window) { reads++; return model.readWindow(window); } }, { readLimits: { maxCells: 10 } });
    const client = await connectWorkbookReads(ports.port2);
    try {
      await expect(client.cells.read({ ...request(model), documentId: 'another-private-workbook' }, signal())).rejects.toMatchObject({ code: 'RevisionConflict' });
      await expect(client.cells.read({ ...request(model), revision: 1 }, signal())).rejects.toMatchObject({ code: 'RevisionConflict' });
      await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
      await expect(client.cells.read({ ...request(model), maxCells: 100_000 }, signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
      await expect(client.cells.read({ ...request(model), unexpected: 'value' } as ReturnType<typeof request>, signal())).rejects.toMatchObject({ code: 'InvalidOperation' });
      expect(reads).toBe(0);
      expect((await client.cells.read({ ...request(model, 10), windows: [] }, signal())).cells).toEqual([]);
    } finally { client.dispose(); server.dispose(); }
  });

  it('cancels promptly, ignores old responses and bounds pending reads and deadlines', async () => {
    const model = empty(), ports = new MessageChannel(), held: number[] = [];
    const reply = (id: number, value: unknown) => ports.port1.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id, ok: true, value });
    ports.port1.onmessage = event => {
      if (event.data.operation === 'snapshot') reply(event.data.id, model.snapshot());
      if (event.data.operation === 'cells') held.push(event.data.id);
    };
    const client = await connectWorkbookReads(ports.port2, { maxPendingRequests: 1, requestTimeoutMs: 100 });
    try {
      const abort = new AbortController(), first = client.cells.read(request(model), abort.signal); await tick(); abort.abort();
      await expect(first).rejects.toMatchObject({ code: 'OperationCancelled' });
      const second = client.cells.read(request(model), signal()); await tick();
      await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: 'ResourceLimitExceeded' });
      reply(held[0]!, { documentId: 'other', revision: 999, cells: [] });
      reply(held[1]!, { documentId: 'transport', revision: 0, cells: [] });
      expect((await second).cells).toEqual([]);
      await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: 'OperationCancelled' });
      const last = client.cells.read(request(model), signal()); client.dispose();
      await expect(last).rejects.toMatchObject({ code: 'OperationCancelled' });
      await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: 'InvalidOperation' });
    } finally { client.dispose(); ports.port1.close(); }
  });

  it('rejects mismatched replies and malformed errors without disclosing raw provider exceptions', async () => {
    const model = empty(), ports = new MessageChannel();
    const server = serveWorkbookReads(ports.port1, { ...model, readWindow() { throw Error('Private database credential'); } });
    const client = await connectWorkbookReads(ports.port2);
    try { await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: 'InvalidOperation', message: 'The authorized workbook transport could not complete the read.' }); }
    finally { client.dispose(); server.dispose(); }
    for (const malformed of [false, true]) {
      const ports = new MessageChannel();
      ports.port1.onmessage = event => {
        if (event.data.operation === 'snapshot') ports.port1.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id: event.data.id, ok: true, value: model.snapshot() });
        if (event.data.operation === 'cells') ports.port1.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, id: event.data.id, ...(malformed
          ? { ok: false, error: { code: 'invented', message: 'Private provider failure' } } : { ok: true, value: { documentId: 'different', revision: 0, cells: [] } }) });
      };
      const client = await connectWorkbookReads(ports.port2);
      try { await expect(client.cells.read(request(model), signal())).rejects.toMatchObject({ code: malformed ? 'InvalidDocument' : 'RevisionConflict' }); }
      finally { client.dispose(); ports.port1.close(); }
    }
  });

  it('bounds and freezes metadata separately and closes incomplete handshakes', async () => {
    const model = empty(), ports = new MessageChannel();
    expect(() => serveWorkbookReads(ports.port1, model, { maxSnapshotNodes: 1 })).toThrow(/budget/); ports.port1.close(); ports.port2.close();
    for (const options of [{ maxSnapshotCodeUnits: 1 }, { maxSnapshotNodes: 1 }]) {
      const ports = new MessageChannel(), server = serveWorkbookReads(ports.port1, model);
      await expect(connectWorkbookReads(ports.port2, options)).rejects.toMatchObject({ code: 'ResourceLimitExceeded' }); server.dispose();
    }
    const waiting = new MessageChannel(), abort = new AbortController(), connection = connectWorkbookReads(waiting.port2, {}, abort.signal); abort.abort();
    await expect(connection).rejects.toMatchObject({ code: 'OperationCancelled' }); waiting.port1.close();
    const closed = new MessageChannel(), closing = connectWorkbookReads(closed.port2);
    closed.port1.postMessage({ protocol: WORKBOOK_READ_PROTOCOL, event: 'close' });
    await expect(closing).rejects.toMatchObject({ code: 'OperationCancelled' }); closed.port1.close();
  });

  it('server shutdown promptly cancels client reads and one connection cannot close another', async () => {
    const model = empty(), a = new MessageChannel(), b = new MessageChannel();
    const firstServer = serveWorkbookReads(a.port1, model), secondServer = serveWorkbookReads(b.port1, model);
    const first = await connectWorkbookReads(a.port2), second = await connectWorkbookReads(b.port2);
    try {
      const pending = first.cells.read(request(model), signal()); firstServer.dispose();
      await expect(pending).rejects.toMatchObject({ code: 'OperationCancelled' });
      expect((await second.cells.read(request(model), signal())).cells).toEqual([]);
      first.dispose(); expect((await second.cells.read(request(model), signal())).cells).toEqual([]);
    } finally { first.dispose(); second.dispose(); firstServer.dispose(); secondServer.dispose(); }
  });
  it('delivers cancellation to real server window batches at their cooperative task boundary', async () => {
    const model = empty(), ports = new MessageChannel(), abort = new AbortController(); let reads = 0;
    const server = serveWorkbookReads(ports.port1, { ...model, readWindow(window) {
      if (++reads === 1) abort.abort(); return model.readWindow(window);
    } });
    const client = await connectWorkbookReads(ports.port2);
    try {
      const batch = { ...request(model, 5000), windows: [{ sheetId: 'sheet-1', row: 0, column: 0, rows: 250, columns: 20 }] };
      await expect(client.cells.read(batch, abort.signal)).rejects.toMatchObject({ code: 'OperationCancelled' });
      await tick(); expect(reads).toBeGreaterThan(0); expect(reads).toBeLessThan(21);
      expect((await client.cells.read(request(model), signal())).cells).toEqual([]);
    } finally { client.dispose(); server.dispose(); }
  });
  it('delivers message cancellation during actual rule evaluation and keeps later reads isolated', async () => {
    const files = new Map(createWorkbook({ documentId: 'transport' })), metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!));
    metadata.workbook.sheets[0].conditionalRules = [{ ruleId: 'rule', kind: 'expression', expression: '=A1=0', priority: 1, stopIfTrue: false,
      origin: { row: 0, column: 0 }, ranges: [{ from: { row: 0, column: 0 }, to: { row: 999, column: 0 } }] }];
    files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
    const model = openDocument(files, { kind: 'workbook' }), ports = new MessageChannel(), abort = new AbortController(); let reads = 0;
    const server = serveWorkbookReads(ports.port1, { ...model, getCell(sheetId, row, column) { if (++reads === 1) abort.abort(); return model.getCell(sheetId, row, column); } });
    const client = await connectWorkbookReads(ports.port2);
    const input = { documentId: 'transport', revision: 0, sheetId: 'sheet-1', coordinates: Array.from({ length: 1000 }, (_, row) => ({ row, column: 0 })), ...DEFAULT_WORKBOOK_RULE_LIMITS };
    try {
      await expect(client.rules!.read(input, abort.signal)).rejects.toMatchObject({ code: 'OperationCancelled' });
      await tick(); expect(reads).toBeGreaterThan(0); expect(reads).toBeLessThan(50);
      expect((await client.rules!.read({ ...input, coordinates: [{ row: 0, column: 0 }] }, signal())).cells[0]!.conditional[0]!.state).toBe('match');
    } finally { client.dispose(); server.dispose(); }
  });
});
