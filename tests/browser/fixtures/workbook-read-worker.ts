import { DocumentError, openDocument } from '@doffi/core';
import type { WorkbookModel } from '@doffi/core';
import { importXlsx } from '@doffi/office';
import { serveWorkbookReads } from '@doffi/render';
import type { WorkbookReadServer } from '@doffi/render';

const scope = globalThis as unknown as { onmessage: ((event: MessageEvent<{ kind: string; documentId: string; bytes: Uint8Array; port: MessagePort }>) => void) | null; postMessage(value: unknown): void };
let model: WorkbookModel | undefined, server: WorkbookReadServer | undefined, loading = false;
scope.onmessage = event => {
  const request = event.data;
  if (request.kind === 'dispose') { server?.dispose(); server = undefined; model = undefined; scope.postMessage({ kind: 'disposed' }); return; }
  if (request.kind === 'audit' && model) {
    // Independent acceptance scan runs after the first screen. It does not
    // add a full-document transfer or artificial work to the mount timing.
    const current = model;
    void (async () => {
      let cells = 0, formulas = 0;
      for (const chunk of current.snapshot().manifest.chunks) {
        for (let row = chunk.row; row < Math.min(chunk.row + 256, 1_048_576); row += 64) {
          const window = current.readWindow({ sheetId: chunk.sheetId, row, column: chunk.column, rows: Math.min(64, 1_048_576 - row), columns: Math.min(32, 16_384 - chunk.column) });
          cells += window.length; formulas += window.filter(cell => cell.formula).length;
        }
        await new Promise<void>(resolve => setTimeout(resolve, 0));
      }
      scope.postMessage({ kind: 'audit', cells, formulas, chunks: current.snapshot().manifest.chunks.length });
    })();
    return;
  }
  if (request.kind !== 'open' || loading || model) return;
  loading = true;
  void (async () => {
    try {
      const files = await importXlsx(request.bytes, { documentId: request.documentId, onProgress(progress) { scope.postMessage({ kind: 'progress', ...progress }); } });
      model = openDocument(files, { kind: 'workbook' });
      server = serveWorkbookReads(request.port, { ...model, serialize() { throw Error('The read server must not transfer a complete package'); } });
    } catch (reason) {
      const error = reason instanceof DocumentError ? reason : new DocumentError('InvalidDocument', 'The native workbook could not be opened.');
      scope.postMessage({ kind: 'error', error: error.toJSON() }); request.port.close();
    } finally { loading = false; }
  })();
};
