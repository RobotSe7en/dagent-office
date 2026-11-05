import { connectWorkbookReads, mountDocumentPreview } from '@dagent-office/render';
import type { WorkbookPreview, WorkbookReadConnection } from '@dagent-office/render';

export async function nativeWorkerRead(code: string, input: number[], documentId: string, audit: boolean) {
  const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), channel = new MessageChannel();
  const container = document.createElement('div'); document.body.append(container);
  let client: WorkbookReadConnection | undefined, preview: WorkbookPreview | undefined, previous = performance.now(), ticks = 0, replies = 0, largestReplyUnits = 0, progress = 0;
  const delays: number[] = [], failures: string[] = [], abort = new AbortController();
  const started = performance.now(), timer = setInterval(() => { const now = performance.now(); delays.push(Math.max(0, now - previous - 10)); previous = now; ticks++; }, 10);
  const waiters = new Map<string, (value: Record<string, unknown>) => void>();
  const message = (event: MessageEvent<Record<string, unknown>>) => {
    if (event.data.kind === 'progress') progress++;
    else if (event.data.kind === 'error') { failures.push(JSON.stringify(event.data.error)); abort.abort(); }
    else waiters.get(String(event.data.kind))?.(event.data);
  };
  worker.addEventListener('message', message); worker.onerror = () => { failures.push('Worker failed'); abort.abort(); };
  channel.port2.addEventListener('message', event => { replies++; largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(event.data).length); });
  const waitFor = (kind: string): Promise<Record<string, unknown>> => new Promise((resolve, reject) => {
    const timeout = setTimeout(() => { waiters.delete(kind); reject(Error('Worker host response timed out')); }, 60_000);
    waiters.set(kind, value => { clearTimeout(timeout); waiters.delete(kind); resolve(value); });
  });
  try {
    const bytes = new Uint8Array(input); worker.postMessage({ kind: 'open', documentId, bytes, port: channel.port1 }, [bytes.buffer, channel.port1]);
    client = await connectWorkbookReads(channel.port2, {}, abort.signal);
    const openedMs = performance.now() - started, importDelays = [...delays].sort((a, b) => a - b), snapshot = client.cells.snapshot;
    preview = mountDocumentPreview(container, client, { widthPt: 420, heightPt: 300 }); await preview.ready;
    const firstScreenMs = performance.now() - started, first = preview.layout();
    const query = { documentId, revision: snapshot.manifest.revision, windows: [{ sheetId: snapshot.metadata.workbook.activeSheetId, row: audit ? 9999 : 0, column: 0, rows: 1, columns: 20 }], maxCells: 20, maxCodeUnits: 65536 };
    const last = await client.cells.read(query, new AbortController().signal);
    if (audit) await preview.setViewport({ ...first.viewport, scrollLeftPt: 1000, scrollTopPt: 149_800 });
    const afterScroll = preview.layout(); preview.dispose(); preview = undefined;
    const retained = await client.cells.read(query, new AbortController().signal);
    let counts: Record<string, unknown> | undefined;
    if (audit) { const result = waitFor('audit'); worker.postMessage({ kind: 'audit' }); counts = await result; }
    const stopped = waitFor('disposed'); worker.postMessage({ kind: 'dispose' }); await stopped;
    let closed = ''; try { await client.cells.read(query, new AbortController().signal); closed = 'accepted'; } catch (error) { closed = (error as { code: string }).code; }
    return { openedMs, firstScreenMs, first, afterScroll, last: last.cells, retained: JSON.stringify(last) === JSON.stringify(retained), closed, counts,
      ticks, progress, largestReplyUnits, replies, importDelayP95Ms: importDelays[Math.floor(importDelays.length * .95)] ?? 0, failures,
      snapshotUnits: JSON.stringify(snapshot).length, frozen: Object.isFrozen(snapshot.metadata.workbook.sheets[0]), detachedInput: bytes.byteLength === 0,
      browser: navigator.userAgent };
  } finally { clearInterval(timer); preview?.dispose(); client?.dispose(); channel.port2.close(); worker.terminate(); URL.revokeObjectURL(url); container.remove(); }
}

export async function cancelNativeWorker(code: string, input: number[]) {
  const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), ports = new MessageChannel(), abort = new AbortController();
  let progress = 0;
  worker.onmessage = event => { if (event.data.kind === 'progress') { progress++; abort.abort(); worker.terminate(); } };
  try {
    const bytes = new Uint8Array(input); worker.postMessage({ kind: 'open', documentId: 'cancel-native', bytes, port: ports.port1 }, [bytes.buffer, ports.port1]);
    let error = ''; try { const connection = await connectWorkbookReads(ports.port2, {}, abort.signal); connection.dispose(); error = 'accepted'; } catch (reason) { error = (reason as { code: string }).code; }
    return { error, progress };
  } finally { abort.abort(); worker.terminate(); ports.port2.close(); URL.revokeObjectURL(url); }
}
Object.assign(window, { workbookReadHost: { nativeWorkerRead, cancelNativeWorker } });
