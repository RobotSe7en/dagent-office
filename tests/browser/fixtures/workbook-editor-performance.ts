import type { WorkbookSnapshot } from '@doffi/core';
import { connectWorkbookSession, mountWorkbookEditor } from '@doffi/editor';
import type { WorkbookRemoteEditorView, WorkbookSessionConnection } from '@doffi/editor';
import { decodeTestBytes } from './package-transport.js';

/** Observe the test's closed surfaces without opening production shadows or
 * retaining every prior window. The active and preparing surfaces are bounded. */
function observeSurfaces() {
  const parents: ShadowRoot[] = [], frames = new Map<HTMLIFrameElement, { roots: ShadowRoot[]; release(): void }>();
  const attach = Element.prototype.attachShadow;
  const observer = new MutationObserver(scan);
  Element.prototype.attachShadow = function(init) { const root = attach.call(this, init); parents.push(root); observer.observe(root, { childList: true, subtree: true }); return root; };
  function scan() {
    for (const [frame, record] of frames) if (!frame.isConnected) { record.release(); frames.delete(frame); }
    for (const parent of parents) for (const frame of parent.querySelectorAll('iframe')) {
      if (frames.has(frame)) continue;
      let restore = () => {}; const roots: ShadowRoot[] = [];
      const loaded = () => {
        const prototype = (frame.contentWindow as Window & typeof globalThis).Element.prototype, original = prototype.attachShadow;
        prototype.attachShadow = function(init) {
          const root = original.call(this, init);
          if (this.localName === 'office-workbook-view-surface') {
            for (let index = roots.length - 1; index >= 0; index--) if (!roots[index]!.host.isConnected) roots.splice(index, 1);
            roots.push(root);
          }
          return root;
        };
        restore = () => { prototype.attachShadow = original; };
      };
      frame.addEventListener('load', loaded, { capture: true, once: true });
      frames.set(frame, { roots, release() { frame.removeEventListener('load', loaded, true); restore(); roots.length = 0; } });
    }
  }
  return {
    frame() { scan(); const active = [...frames.keys()].filter(frame => frame.isConnected); if (active.length !== 1) throw Error('Expected one accepted editor iframe'); return active[0]!; },
    cell(row: number, column: number) {
      scan();
      for (const { roots } of frames.values()) for (const root of roots) if (root.host.isConnected) {
        const cell = root.querySelector(`[data-office-view="workbook-cell"][data-row="${row}"][data-column="${column}"]`);
        if (cell) return cell.textContent?.trim();
      }
      return undefined;
    },
    input() { const input = parents.flatMap(root => [...root.querySelectorAll('textarea')])[0]; if (!input) throw Error('Editor input is missing'); return input; },
    button(label: string) { const button = parents.flatMap(root => [...root.querySelectorAll('button')]).find(button => button.textContent === label); if (!button || button.disabled) throw Error('Editor control is unavailable: ' + label); return button; },
    dispose() { observer.disconnect(); Element.prototype.attachShadow = attach; for (const record of frames.values()) record.release(); frames.clear(); parents.length = 0; },
  };
}

/** Actual mounted editor and visible dependent formula. This diagnostic does
 * not certify the fixed-resource stage's CPU, memory or latency limits. */
export async function ordinaryMountedWorkerEdits(code: string, encoded: string) {
  const documentId = 'ordinary-mounted-editor', url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), channel = new MessageChannel();
  const host = document.createElement('div'); document.body.append(host);
  let client: WorkbookSessionConnection | undefined, editor: WorkbookRemoteEditorView | undefined, snapshot: WorkbookSnapshot | undefined;
  let captureRows = false, maximumReadRow = -1, largestReplyUnits = 0, sameFrame = true, ticks = 0;
  const failures: unknown[] = [], samples: { inputMs: number; previewMs: number; totalMs: number }[] = [], surfaces = observeSurfaces();
  const timer = setInterval(() => ticks++, 10); let auditResolve!: (value: Record<string, unknown>) => void;
  const audited = new Promise<Record<string, unknown>>(resolve => { auditResolve = resolve; });
  worker.onmessage = event => { if (event.data.kind === 'audit') auditResolve(event.data); if (event.data.kind === 'error') failures.push(event.data.error); };
  worker.onerror = () => { failures.push('Worker failure'); };
  channel.port2.addEventListener('message', event => {
    largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(event.data).length);
    if (event.data.port instanceof MessagePort) event.data.port.addEventListener('message', (reply: MessageEvent<{ value?: { cells?: { row: number }[] } }>) => {
      largestReplyUnits = Math.max(largestReplyUnits, JSON.stringify(reply.data).length);
      if (captureRows && Array.isArray(reply.data.value?.cells)) for (const cell of reply.data.value.cells) maximumReadRow = Math.max(maximumReadRow, cell.row);
    });
  });
  try {
    const start = performance.now(), bytes = decodeTestBytes(encoded);
    worker.postMessage({ kind: 'open', documentId, bytes, port: channel.port1 }, [bytes.buffer, channel.port1]);
    client = await connectWorkbookSession(channel.port2);
    editor = mountWorkbookEditor(host, client, { widthPt: 420, heightPt: 240,
      prepareOptions: () => ({ calculationContext: { now: { unixMilliseconds: Date.now(), utcOffsetMinutes: -new Date().getTimezoneOffset() } } }),
      onViewChange(value) { snapshot = value; }, onError(error) { failures.push({ message: error.message, code: (error as { code?: string }).code }); } });
    await editor.ready; const firstScreenMs = performance.now() - start, initial = snapshot!, originalFrame = surfaces.frame(), sheetId = initial.metadata.workbook.activeSheetId;
    const at = { row: 0, column: 2 }, dependent = { row: 0, column: 19 };
    // Leave T1 on screen while the formula bar edits C1. Each sample verifies
    // the editor's own rendered batch without a second verification calculation.
    await editor.selectRange(sheetId, { from: dependent, to: dependent }); const viewport = editor.layout().viewport;
    await editor.selectRange(sheetId, { from: at, to: at }); await editor.setViewport(viewport); await editor.settled();
    if (surfaces.cell(0, 19) !== '492') throw Error('Initial dependent formula is not visible and current');
    const input = surfaces.input(), setupMs = performance.now() - start - firstScreenMs; captureRows = true;
    for (let index = 0; index < 100; index++) {
      const before = performance.now(); input.value = String(100 + index); input.dispatchEvent(new Event('input', { bubbles: true, composed: true }));
      const receipt = await editor.applyInput(); if (!receipt?.changed) throw Error('Mounted editor input did not commit'); const committed = performance.now();
      await editor.settled(); const shown = performance.now();
      if (surfaces.cell(0, 19) !== String(570 + index)) throw Error('Mounted editor displayed an incorrect dependent formula');
      sameFrame &&= surfaces.frame() === originalFrame;
      samples.push({ inputMs: committed - before, previewMs: shown - committed, totalMs: shown - before });
    }
    captureRows = false; const fullHistory = client.state().history;
    async function history(label: string) {
      for (let index = 0; index < 100; index++) {
        const revision = client!.state().revision; surfaces.button(label).click(); await editor!.settled();
        if (client!.state().revision !== revision + 1) throw Error('Mounted editor history did not advance exactly one revision');
      }
    }
    await history('撤销'); const undoneHistory = client.state().history, undoneText = surfaces.cell(0, 19);
    await history('重做'); const redoneHistory = client.state().history, redoneText = surfaces.cell(0, 19);
    await history('撤销'); const originalChunksRestored = JSON.stringify(snapshot!.manifest.chunks) === JSON.stringify(initial.manifest.chunks), restoredInput = input.value, restoredText = surfaces.cell(0, 19);
    const finalRevision = snapshot!.manifest.revision; editor.dispose(); editor = undefined; client.dispose(); client = undefined;
    const timeout = setTimeout(() => { failures.push('Audit timeout'); auditResolve({}); }, 15_000);
    worker.postMessage({ kind: 'closed-audit' }); const audit = await audited; clearTimeout(timeout);
    const sorted = samples.map(sample => sample.totalMs).sort((a, b) => a - b);
    return { scope: 'Unrestricted mounted public Worker editor with a visible dependent formula; fixed resource acceptance remains separate.',
      firstScreenMs, setupMs, samples, editAndPreviewP95Ms: sorted[Math.ceil(sorted.length * .95) - 1]!, maximumMs: sorted.at(-1)!,
      maximumReadRow, largestReplyUnits, sameFrame, originalChunksRestored, restoredInput, undoneText, redoneText, restoredText, finalRevision,
      fullHistory, undoneHistory, redoneHistory, audit, ticks, failures, userAgent: navigator.userAgent };
  } finally { clearInterval(timer); editor?.dispose(); client?.dispose(); surfaces.dispose(); channel.port2.close(); worker.terminate(); URL.revokeObjectURL(url); host.remove(); }
}
Object.assign(window, { workbookEditorPerformance: { ordinaryMountedWorkerEdits } });
