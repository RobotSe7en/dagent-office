import { createWorkbookNumberFormatter, DocumentError, openDocument } from '@dagent-office/core';
import type { DocumentSelection, WorkbookDocumentMetadata, WorkbookModel } from '@dagent-office/core';
import { mountDocumentPreview } from '@dagent-office/render';
import type { WorkbookPreview, WorkbookPreviewSource, WorkbookPreviewOptions } from '@dagent-office/render';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { modelRuleSource } from '../../../packages/render/src/workbook/rule-source.js';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import { decorateWorkbookPreview } from '../../../packages/render/src/workbook/preview-decoration.js';
import type { WorkbookPreviewDecoration } from '../../../packages/render/src/workbook/preview-decoration.js';

type Entries = readonly [string, number[]][];
function modelOf(entries: Entries, rtl = false): WorkbookModel {
  const files = new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)]));
  const metadata = JSON.parse(new TextDecoder().decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata;
  metadata.workbook.sheets[0]!.rightToLeft = rtl;
  metadata.workbook.sheets.push({ sheetId: 'preview-empty', name: 'Empty', visibility: 'visible', defaultRowHeightPt: 20, defaultColumnWidth: 8.43,
    defaultCssClass: metadata.workbook.sheets[0]!.defaultCssClass, rowDimensions: [], columnDimensions: [], merges: [], frozen: { rows: 0, columns: 0 }, showGridLines: true, rightToLeft: rtl });
  files.set('metadata.json', new TextEncoder().encode(JSON.stringify(metadata)));
  return openDocument(files, { kind: 'workbook' });
}
function host() { const container = document.createElement('div'); container.style.cssText = 'position:fixed;left:0;top:0;z-index:2147483647;background:white'; document.body.append(container); return container; }
// Capture only our own chrome shadow for inspecting public UI. The production
// component continues to expose neither its source DOM nor its private surface.
function mount(container: HTMLElement, input: WorkbookModel | WorkbookPreviewSource, callbacks: WorkbookPreviewOptions = {}) {
  const original = Element.prototype.attachShadow; let chrome!: ShadowRoot;
  Element.prototype.attachShadow = function (init) { const root = original.call(this, init); chrome ??= root; return root; };
  try { const preview = mountDocumentPreview(container, input, { widthPt: 420, heightPt: 300, ...callbacks }); return { preview, chrome }; }
  finally { Element.prototype.attachShadow = original; }
}
let cleanup = () => {}, inspect = () => ({}), action = async (_action: string) => ({});
export function previewInspect() { return inspect(); }
export function previewAction(name: string) { return action(name); }
export async function mountPublicPreview(entries: Entries, rtl: boolean) {
  cleanup(); const model = modelOf(entries, rtl), before = model.serialize(), container = host(), errors: string[] = [], selections: (DocumentSelection | undefined)[] = [];
  let serializations = 0;
  const direct = { ...model, serialize() { serializations++; throw Error('A preview must not serialize the entire workbook'); } };
  const { preview, chrome } = mount(container, direct, { onError: error => errors.push(error.code), onSelectionChange: selection => selections.push(selection) });
  try {
    await preview.ready;
    const frame = chrome.querySelector('iframe')!, rect = frame.getBoundingClientRect(), sheetId = preview.layout().sheetId;
    const bytesUnchanged = () => { const after = model.serialize(); return [...before].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, i) => byte === after.get(path)![i])); };
    inspect = () => ({ report: preview.layout(), address: chrome.querySelector<HTMLInputElement>('[aria-label="单元格地址"]')!.value,
      value: chrome.querySelector<HTMLInputElement>('[aria-label="单元格值或公式"]')!.value, status: chrome.querySelector('[role="status"]')!.textContent,
      validation: chrome.querySelector('[aria-label="数据验证"]')!.textContent,
      tabs: [...chrome.querySelectorAll('[role="tab"]')].map(tab => ({ text: tab.textContent, selected: tab.getAttribute('aria-selected') })),
      selection: preview.selection(), selections, errors, serializations, unchanged: bytesUnchanged(), frames: chrome.querySelectorAll('iframe').length });
    action = async name => {
      if (name.startsWith('point-')) {
        const [, column, row] = name.split('-'), headers = [...chrome.querySelectorAll('.header>div')];
        const x = headers.find(node => node.textContent === column)!.getBoundingClientRect(), y = headers.find(node => node.textContent === row)!.getBoundingClientRect();
        return { x: x.x + x.width / 2, y: y.y + y.height / 2 };
      }
      if (name === 'empty') chrome.querySelector<HTMLButtonElement>('[data-sheet-id="preview-empty"]')!.click();
      if (name === 'first') await preview.showSheet(sheetId);
      if (name === 'comment') preview.openComment(model.snapshot().annotations[0]!.annotationId);
      if (name === 'close') preview.closeComment();
      if (name === 'bad-sheet') { try { await preview.showSheet('missing'); } catch { /* inspected below */ } return inspect(); }
      if (name === 'dispose') { const fields = [...chrome.querySelectorAll('input')]; preview.dispose(); cleanup(); return { removed: !container.isConnected, frames: chrome.querySelectorAll('iframe').length, unchanged: bytesUnchanged(), serializations, fieldsCleared: fields.every(field => !field.value) }; }
      await preview.settled(); return inspect();
    };
    cleanup = () => { preview.dispose(); container.remove(); cleanup = () => {}; };
    return { x: rect.x + (rtl ? rect.width - 16 : 16), y: rect.y + 8, ...inspect() };
  } catch (error) { preview.dispose(); container.remove(); throw error; }
}

export async function publicPreviewLifecycle(entries: Entries) {
  const model = modelOf(entries), cells = modelWindowSource(model), assets = modelAssetSource(model), charts = modelChartSource(model), container = host();
  let gate: { begin(): void; release(fail: boolean): void; signal: AbortSignal | undefined } | undefined, wrongRevision = false;
  const source: WorkbookPreviewSource = { kind: 'workbook', assets, charts, cells: { ...cells, async read(request, signal) {
    if (gate) { const held = gate; gate = undefined; held.signal = signal; await new Promise<void>((resolve, reject) => { held.release = fail => fail ? reject(Error('Late private failure')) : resolve(); held.begin(); }); }
    const response = await cells.read(request, new AbortController().signal);
    return wrongRevision ? { ...response, revision: response.revision + 1 } : response;
  } } };
  const stall = () => {
    let begin!: () => void; const started = new Promise<void>(resolve => { begin = resolve; });
    const held = { begin, release: (_fail: boolean) => {}, signal: undefined as AbortSignal | undefined }; gate = held;
    return { started, release: (fail = false) => held.release(fail), aborted: () => held.signal?.aborted };
  };
  const errors: string[] = [], unhandled: string[] = [], rejection = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', rejection);
  const created = new Set<string>(), revoked = new Set<string>(), originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL;
  URL.createObjectURL = blob => { const url = originalCreate.call(URL, blob); created.add(url); return url; }; URL.revokeObjectURL = url => { revoked.add(url); originalRevoke.call(URL, url); };
  let preview: WorkbookPreview | undefined;
  try {
    const mounted = mount(container, source, { onError: error => errors.push(error.code) }); preview = mounted.preview; const chrome = mounted.chrome;
    await preview.ready; const first = preview.layout(), oldFrame = chrome.querySelector('iframe')!;
    const scroll = stall(), obsolete = preview.setViewport({ ...first.viewport, scrollTopPt: 30 }); await scroll.started;
    await preview.showSheet('preview-empty'); const newest = await obsolete;
    scroll.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const scrollRace = newest.sheetId === 'preview-empty' && scroll.aborted() && preview.layout().sheetId === newest.sheetId;
    await preview.showSheet(first.sheetId);
    wrongRevision = true; preview.update(source); let failed = '';
    try { await preview.settled(); } catch (error) { failed = (error as { code: string }).code; }
    const retained = preview.layout().sheetId === first.sheetId && chrome.querySelector('iframe') === oldFrame && chrome.querySelectorAll('iframe').length === 1;
    wrongRevision = false;
    const replacement = stall(); preview.update(source); await replacement.started;
    const pending = preview.settled(); preview.update(model); await preview.settled(); await pending;
    replacement.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const updateRace = replacement.aborted() && !oldFrame.isConnected && chrome.querySelectorAll('iframe').length === 1;
    const stopped = stall(); preview.update(source); await stopped.started;
    const last = preview.settled().then(() => 'accepted', error => (error as { code: string }).code); preview.dispose(); const disposed = await last;
    stopped.release(true); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const providerAlive = (await cells.read({ documentId: first.documentId, revision: first.revision, windows: [], maxCells: 1, maxCodeUnits: 1 }, new AbortController().signal)).revision === first.revision;
    return { scrollRace, failed, retained, updateRace, disposed, errors, unhandled, providerAlive, released: [...created].every(url => revoked.has(url)), removed: !mounted.preview.element.isConnected };
  } finally { preview?.dispose(); container.remove(); window.removeEventListener('unhandledrejection', rejection); URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; }
}
export async function previewDecorationMeasurement(entries: Entries) {
  const model = modelOf(entries), snapshot = model.snapshot(), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  let mode = 'ok', released = 0, controls: WorkbookPreviewDecoration | undefined;
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter, {}, (window, content, maxNodes) => {
    if (mode === 'throw') throw new DocumentError('UnsupportedFeature', 'Deliberate control construction failure.');
    if (mode === 'budget') return { nodes: maxNodes + 1, dispose() { released++; } };
    const next = decorateWorkbookPreview(surface, snapshot, window, content, maxNodes, action => action());
    if (mode === 'cancel') controller.cancelPending(); else controls = next;
    return { get nodes() { return next.nodes; }, get fonts() { return next.fonts; }, dispose() { next.dispose(); released++; } };
  });
  const viewport = { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 }, errors: string[] = [];
  try {
    await controller.render(snapshot.metadata.workbook.activeSheetId, viewport);
    const accepted = controller.current!, html = surface.view.innerHTML;
    for (mode of ['throw', 'budget', 'cancel']) {
      try { errors.push(await controller.render(accepted.frame.sheetId, { ...viewport, heightPt: 320 }) ? 'accepted' : 'cancelled'); } catch (error) { errors.push((error as { code: string }).code); }
      if (controller.current !== accepted || surface.view.innerHTML !== html) throw Error('Control failure changed the accepted window');
    }
    controls!.openComment(snapshot.annotations[0]!.annotationId); controls!.selectCell(0, 0);
    const walker = surface.document.createTreeWalker(accepted.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
    const result = { errors, nodes: accepted.nodes, actualNodes, fonts: accepted.fonts.length, panel: accepted.element.querySelector('[data-office-view="workbook-preview-comments"]')?.textContent,
      highlights: accepted.element.querySelector('[data-office-view="workbook-cell-selection"]')?.childNodes.length, releasedCandidates: released };
    controls!.closeComment(); controller.dispose();
    return { ...result, releasedAll: released, empty: !surface.view.childNodes.length && !surface.source.childNodes.length };
  } finally { controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}
export async function initialPreviewFailure(entries: Entries) {
  const model = modelOf(entries), container = host(), errors: string[] = [];
  const missing = mount(container, model, { sheetId: 'missing', onError: error => errors.push(error.code) });
  let rejected = '';
  try { await missing.preview.ready; rejected = 'accepted'; } catch (error) { rejected = (error as { code: string }).code; }
  const frames = missing.chrome.querySelectorAll('iframe').length; missing.preview.dispose();
  const cells = modelWindowSource(model); let fail = true;
  const source: WorkbookPreviewSource = { kind: 'workbook', assets: modelAssetSource(model), charts: modelChartSource(model), cells: { ...cells, async read(request, signal) {
    if (fail) { fail = false; throw Error('Private transport details'); } return cells.read(request, signal);
  } } };
  let preview!: WorkbookPreview;
  const retry = mount(container, source, { onError: error => { errors.push(error.code); preview.update(model); } }); preview = retry.preview;
  try { await preview.ready; const report = await preview.settled(); return { rejected, frames, errors, retried: report.documentId === model.snapshot().manifest.documentId, finalFrames: retry.chrome.querySelectorAll('iframe').length }; }
  finally { preview.dispose(); container.remove(); }
}
export async function ruleWindowLifecycle(entries: Entries) {
  const model = modelOf(entries), snapshot = model.snapshot(), container = host(), rules = modelRuleSource(model), formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
  let mode = 'ok', begin!: () => void, release!: () => void, aborted = false;
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model),
    rules: { ...rules, async read(request, signal) {
      if (mode === 'held') { mode = 'ok'; signal.addEventListener('abort', () => { aborted = true; }, { once: true }); await new Promise<void>(resolve => { release = resolve; begin(); }); }
      const value = await rules.read(request, new AbortController().signal);
      return mode === 'empty' ? { ...value, cells: [] } : mode === 'revision' ? { ...value, revision: value.revision + 1 } : value;
    } } }, formatter);
  const viewport = { widthPt: 420, heightPt: 300, scrollTopPt: 0, scrollLeftPt: 0 }, failures: string[] = [];
  try {
    await controller.render('sheet-1', viewport); const original = controller.current!, before = surface.view.innerHTML;
    for (mode of ['empty', 'revision']) {
      try { await controller.render('sheet-1', viewport); failures.push('accepted'); } catch (error) { failures.push((error as { code: string }).code); }
      if (controller.current !== original || surface.view.innerHTML !== before) throw Error('A failed rule read changed the accepted window');
    }
    mode = 'held'; const started = new Promise<void>(resolve => { begin = resolve; }), old = controller.render('sheet-1', viewport); await started;
    await controller.render('preview-empty', viewport); const cancelled = !await old;
    release(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    return { failures, cancelled, aborted, retainedLatest: controller.current!.frame.sheetId === 'preview-empty' };
  } finally { controller.dispose(); surface.dispose(); formatter.dispose(); container.remove(); }
}
(window as unknown as { workbookPreview: unknown }).workbookPreview = { mountPublicPreview, previewInspect, previewAction, publicPreviewLifecycle, previewDecorationMeasurement, initialPreviewFailure, ruleWindowLifecycle, publicReadAhead };

export async function publicReadAhead(entries: Entries, allow: boolean, rows: number, mode: 'normal' | 'hold' | 'failure' | 'revision' | 'active' | 'preference' = 'normal') {
  const model = modelOf(entries), before = model.serialize(), snapshot = model.snapshot(), source = modelWindowSource(model), container = host();
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 240, heightPt: 60 }), formatter = createWorkbookNumberFormatter({ dateSystem: '1900' });
  let observing = false, foreground = false, reads = 0, largestSlots = 0, held: { signal: AbortSignal; release(): void } | undefined;
  let began!: () => void; const started = new Promise<void>(resolve => { began = resolve; });
  const unhandled: string[] = [], onUnhandled = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', onUnhandled);
  const cells = { ...source, get allowReadAhead() { if (mode === 'preference' && foreground) throw Error('Optional preference failed'); return allow; }, async read(request: Parameters<typeof source.read>[0], signal: AbortSignal) {
    const background = observing && !foreground;
    if (background) {
      reads++; largestSlots = Math.max(largestSlots, request.windows.reduce((sum, item) => sum + item.rows * item.columns, 0)); began();
      if (mode === 'hold' && !held) await new Promise<void>(resolve => { held = { signal, release: resolve }; });
      if (mode === 'failure') throw Error('Private optional read failure');
    }
    const response = await source.read(request, held?.signal === signal ? new AbortController().signal : signal);
    if (background && mode === 'revision') return { ...response, revision: response.revision + 1 };
    if (background && mode === 'active' && response.cells.length) return { ...response, cells: response.cells.map((cell, index) => index ? cell : { ...cell, html: cell.html.replace('<td', '<td onclick="window.readAheadExecuted=true"') }) };
    return response;
  } };
  const providers = { cells, assets: modelAssetSource(model), charts: modelChartSource(model) }, invalid: string[] = [];
  for (const readAheadRows of [-1, 1.5, 65]) try { new WorkbookWindowController(surface, providers, formatter, { readAheadRows }); invalid.push('accepted'); } catch (error) { invalid.push((error as { code: string }).code); }
  const controller = new WorkbookWindowController(surface, providers, formatter, { readAheadRows: rows });
  const pause = (ms: number) => new Promise<void>(resolve => setTimeout(resolve, ms));
  try {
    const viewport = { widthPt: 240, heightPt: 60, scrollLeftPt: 0, scrollTopPt: 0 };
    await controller.render('sheet-1', viewport); observing = true;
    const first = controller.current!, original = first.element.innerHTML;
    let aborted = false, whileHeld = 0;
    if (allow && rows && mode === 'hold') {
      await Promise.race([started, pause(2000).then(() => { throw Error('Read ahead did not start'); })]);
      surface.document.dispatchEvent(new surface.document.defaultView!.PointerEvent('pointerdown', { bubbles: true }));
      foreground = true; await controller.render('sheet-1', { ...viewport, scrollTopPt: 30 }); foreground = false;
      aborted = held!.signal.aborted; await pause(80); whileHeld = reads;
      held!.release();
    }
    await pause(300);
    const accepted = controller.current!, html = accepted.element.innerHTML, optionalReads = reads;
    foreground = true; await controller.render('sheet-1', { ...viewport, scrollTopPt: 60 }); foreground = false;
    const last = controller.current!; controller.dispose(); await pause(80);
    const stopped = reads; await pause(100);
    return { invalid, reads: optionalReads, largestSlots, aborted, whileHeld, unhandled,
      firstRetained: mode === 'hold' ? accepted !== first : accepted === first && html === original,
      foreground: last.frame.layout.scrollTopPt === 60, stopped: stopped === reads,
      clean: !surface.view.childNodes.length && !surface.source.childNodes.length,
      activeContent: Boolean((surface.document.defaultView as unknown as { readAheadExecuted?: boolean }).readAheadExecuted),
      unchanged: [...before].every(([path, bytes]) => { const after = model.serialize().get(path)!; return bytes.length === after.length && bytes.every((byte, index) => byte === after[index]); }) };
  } finally { held?.release(); controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); window.removeEventListener('unhandledrejection', onUnhandled); }
}
