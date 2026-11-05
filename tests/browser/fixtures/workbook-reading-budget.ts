import { connectWorkbookReads, mountDocumentPreview } from '@dagent-office/render';
import type { WorkbookPreview, WorkbookReadConnection } from '@dagent-office/render';
import { cancelNativeWorker } from './workbook-read-host.js';

let stop = () => {}, state = () => ({}), point = (_index: number) => ({ x: 0, y: 0 });
let sample = async (_kind: 'selection' | 'wheel', _index: number) => 0;
let audit = async () => ({});
let cancel = async () => ({});
export function budgetState() { return state(); }
export function budgetPoint(index: number) { return point(index); }
export function budgetSample(kind: 'selection' | 'wheel', index: number) { return sample(kind, index); }
export function budgetAudit() { return audit(); }
export function budgetCancel() { return cancel(); }
export function closeBudget() { stop(); }
export async function openBudget(code: string, input: number[]) {
  stop(); document.body.style.margin = '0';
  const container = document.createElement('div'); document.body.replaceChildren(container);
  const url = URL.createObjectURL(new Blob([code], { type: 'text/javascript' })), worker = new Worker(url), ports = new MessageChannel(), abort = new AbortController();
  const roots: ShadowRoot[] = [], restore: (() => void)[] = [], attach = Element.prototype.attachShadow;
  Element.prototype.attachShadow = function(init) { const root = attach.call(this, init); roots.push(root); observer.observe(root, { childList: true, subtree: true }); return root; };
  restore.push(() => { Element.prototype.attachShadow = attach; });
  const observer = new MutationObserver(records => {
    for (const record of records) for (const node of record.addedNodes) if (node instanceof HTMLIFrameElement) node.addEventListener('load', () => {
      const view = node.contentWindow; if (!view) return;
      const prototype = (view as Window & typeof globalThis).Element.prototype, original = prototype.attachShadow;
      prototype.attachShadow = function(init) { const root = original.call(this, init); roots.push(root); return root; };
      restore.push(() => { prototype.attachShadow = original; });
    }, { capture: true, once: true });
  });
  observer.observe(container, { childList: true, subtree: true });
  let connection: WorkbookReadConnection | undefined, preview: WorkbookPreview | undefined;
  let selectionStart: number | undefined, wheelStart = 0, largestReply = 0, maxFragments = 0;
  const selections: number[] = [], wheels: number[] = [], errors: unknown[] = [], delays: number[] = [], inputDelays: number[] = [];
  const pending = new Map<string, { resolve(value: number): void; reject(error: Error): void; timer: ReturnType<typeof setTimeout> }>();
  let previous = performance.now(); const timer = setInterval(() => { const now = performance.now(); delays.push(Math.max(0, now - previous - 10)); previous = now; }, 10);
  const frame = () => roots.flatMap(root => [...root.querySelectorAll('iframe')])[0]!;
  const grid = () => roots.find(root => root.querySelector('[role="grid"]'))?.querySelector<HTMLElement>('[role="grid"]');
  const count = () => {
    const actual = roots.reduce((sum, root) => sum + root.querySelectorAll('[data-office-view="workbook-cell"],[data-office-view="workbook-text-overflow"]').length, 0);
    maxFragments = Math.max(maxFragments, actual, preview!.layout().cellFragmentCount); return actual;
  };
  const publish = (kind: 'selection' | 'wheel', value: number) => {
    const values = kind === 'selection' ? selections : wheels, index = values.push(value) - 1, waiter = pending.get(kind + index);
    if (waiter) { clearTimeout(waiter.timer); pending.delete(kind + index); waiter.resolve(value); } count();
  };
  ports.port2.addEventListener('message', event => { largestReply = Math.max(largestReply, JSON.stringify(event.data).length); });
  worker.addEventListener('error', event => { errors.push(event.message); abort.abort(); });
  worker.addEventListener('message', event => { if (event.data.kind === 'error') { errors.push(event.data.error); abort.abort(); } });
  stop = () => {
    clearInterval(timer); observer.disconnect(); for (const reset of restore.reverse()) reset();
    for (const waiter of pending.values()) { clearTimeout(waiter.timer); waiter.reject(Error('Budget fixture closed')); } pending.clear();
    preview?.dispose(); connection?.dispose(); abort.abort(); worker.terminate(); ports.port2.close(); URL.revokeObjectURL(url); container.remove();
  };
  try {
    const started = performance.now(), bytes = new Uint8Array(input);
    worker.postMessage({ kind: 'open', documentId: 'reading-budget', bytes, port: ports.port1 }, [bytes.buffer, ports.port1]);
    connection = await connectWorkbookReads(ports.port2, {}, abort.signal); const openedMs = performance.now() - started;
    preview = mountDocumentPreview(container, connection, { widthPt: 880, heightPt: 450,
      onError(error) { errors.push(error); }, onSelectionChange(selection) {
        if (!selection || selectionStart === undefined) return;
        const start = selectionStart; selectionStart = undefined;
        requestAnimationFrame(() => publish('selection', performance.now() - start));
      },
    });
    await preview.ready;
    const firstScreenMs = performance.now() - started, first = preview.layout(), initialFragments = count();
    const document = frame().contentDocument!;
    const inputStart = (event: Event): number => {
      const now = performance.now(), delay = document.defaultView!.performance.now() - event.timeStamp;
      if (!event.isTrusted || delay < -2 || delay > 60_000) errors.push('The native input timestamp could not be calibrated.');
      inputDelays.push(delay); return now - Math.max(0, delay);
    };
    document.addEventListener('pointerdown', event => { selectionStart = inputStart(event); }, true);
    document.addEventListener('wheel', event => {
      wheelStart = inputStart(event); const start = wheelStart, before = preview!.layout().viewport;
      // A microtask checkpoint can run between capture/bubble listeners. A
      // task waits for the SDK wheel handler to register its actual candidate.
      setTimeout(() => { void preview!.settled().then(report => {
        if (report.viewport.scrollTopPt === before.scrollTopPt && report.viewport.scrollLeftPt === before.scrollLeftPt) throw Error('The native wheel did not publish a changed viewport.');
        return new Promise<void>(resolve => requestAnimationFrame(() => resolve()));
      }).then(() => publish('wheel', performance.now() - start)).catch(error => errors.push(String(error))); }, 0);
    }, { capture: true, passive: true });
    point = index => {
      const cell = grid()!.querySelectorAll<HTMLElement>('[role="gridcell"]')[index % 5]!, local = cell.getBoundingClientRect(), outer = frame().getBoundingClientRect();
      return { x: outer.x + local.x + local.width / 2, y: outer.y + local.y + local.height / 2 };
    };
    sample = (kind, index) => {
      const values = kind === 'selection' ? selections : wheels; if (values[index] !== undefined) return Promise.resolve(values[index]!);
      return new Promise((resolve, reject) => { const timer = setTimeout(() => { pending.delete(kind + index); reject(Error('Interaction sample timed out')); }, 15_000); pending.set(kind + index, { resolve, reject, timer }); });
    };
    state = () => ({ openedMs, firstScreenMs, first, initialFragments, maxFragments, largestReply, selections: [...selections], wheels: [...wheels], errors: [...errors], delays: [...delays], inputDelays: [...inputDelays], browser: navigator.userAgent });
    audit = () => new Promise((resolve, reject) => {
      const timer = setTimeout(() => { worker.removeEventListener('message', listener); reject(Error('Audit timed out')); }, 60_000);
      const listener = (event: MessageEvent) => { if (event.data.kind === 'audit') { clearTimeout(timer); worker.removeEventListener('message', listener); resolve(event.data); } };
      worker.addEventListener('message', listener); worker.postMessage({ kind: 'audit' });
    });
    cancel = async () => {
      const before = JSON.stringify(preview!.layout()), result = await cancelNativeWorker(code, input);
      return { ...result, retained: JSON.stringify(preview!.layout()) === before };
    };
    return state();
  } catch (error) { stop(); throw error; }
}
Object.assign(window, { workbookBudget: { openBudget, budgetState, budgetPoint, budgetSample, budgetAudit, budgetCancel, closeBudget } });
