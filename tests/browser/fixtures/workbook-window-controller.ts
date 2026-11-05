import { createWorkbookNumberFormatter, openDocument } from '@dagent-office/core';
import type { WorkbookDocumentMetadata } from '@dagent-office/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';

export async function windowControllerMeasurement(entries: readonly [string, number[]][]) {
  const files = new Map(entries.map(([path, bytes]) => [path, new Uint8Array(bytes)])), decoder = new TextDecoder(), encoder = new TextEncoder();
  const metadata = JSON.parse(decoder.decode(files.get('metadata.json')!)) as WorkbookDocumentMetadata, sheet = metadata.workbook.sheets[0]!;
  sheet.frozen = { rows: 2, columns: 1 };
  for (const object of metadata.workbook.objects ?? []) object.anchor = { kind: 'absolute', xPt: 12, yPt: 12, widthPt: object.kind === 'chart' ? 360 : 220, heightPt: object.kind === 'chart' ? 220 : 150 };
  metadata.workbook.sheets.push({ sheetId: 'controller-empty', name: 'Empty', visibility: 'visible', defaultRowHeightPt: 20, defaultColumnWidth: 8.43,
    defaultCssClass: sheet.defaultCssClass, rowDimensions: [], columnDimensions: [], merges: [], frozen: { rows: 0, columns: 0 }, showGridLines: true, rightToLeft: true });
  files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize();
  const container = document.createElement('div'); container.style.position = 'relative'; document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
  const formatter = createWorkbookNumberFormatter({ dateSystem: metadata.workbook.dateSystem }), cells = modelWindowSource(model), assets = modelAssetSource(model), charts = modelChartSource(model);
  const created = new Set<string>(), revoked = new Set<string>(), originalCreate = URL.createObjectURL, originalRevoke = URL.revokeObjectURL;
  URL.createObjectURL = blob => { const url = originalCreate.call(URL, blob); created.add(url); return url; };
  URL.revokeObjectURL = url => { revoked.add(url); originalRevoke.call(URL, url); };
  const unhandled: string[] = [], rejection = (event: PromiseRejectionEvent) => { unhandled.push(String(event.reason)); event.preventDefault(); };
  window.addEventListener('unhandledrejection', rejection);
  type Phase = 'cells' | 'assets' | 'charts';
  let next: { phase: Phase; hold(signal: AbortSignal, read: (signal: AbortSignal) => Promise<unknown>): Promise<unknown> } | undefined;
  const intercept = <T>(phase: Phase, signal: AbortSignal, read: (signal: AbortSignal) => Promise<T>): Promise<T> => {
    if (next?.phase !== phase) return read(signal);
    const gate = next; next = undefined; return gate.hold(signal, read) as Promise<T>;
  };
  const stall = (phase: Phase) => {
    let begin!: () => void, release!: (fail?: boolean) => Promise<void>, signal: AbortSignal;
    const started = new Promise<void>(resolve => { begin = resolve; });
    next = { phase, hold(input, read) {
      signal = input;
      return new Promise((resolve, reject) => {
        release = async (fail = false) => { if (fail) reject(Error('Late provider failure')); else resolve(await read(new AbortController().signal)); };
        begin();
      });
    } };
    return { started, release: (fail?: boolean) => release(fail), aborted: () => signal.aborted };
  };
  let wrongRevision = false, brokenObject = false;
  const providers = {
    cells: { ...cells, async read(request: Parameters<typeof cells.read>[0], signal: AbortSignal) {
      const response = await intercept('cells', signal, signal => cells.read(request, signal));
      if (!wrongRevision) return response; wrongRevision = false; return { ...response, revision: response.revision + 1 };
    } },
    assets: { ...assets, async readAsset(request: Parameters<typeof assets.readAsset>[0], signal: AbortSignal) {
      const response = await intercept('assets', signal, signal => assets.readAsset(request, signal));
      if (!brokenObject) return response; brokenObject = false; return { ...response, bytes: new Uint8Array([1]) };
    } },
    charts: { ...charts, async read(request: Parameters<typeof charts.read>[0], signal: AbortSignal) {
      const response = await intercept('charts', signal, signal => charts.read(request, signal));
      if (!brokenObject) return response; brokenObject = false; return { ...response, charts: [] };
    } },
  };
  const controller = new WorkbookWindowController(surface, providers, formatter), viewport = { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 };
  const errors: string[] = [], races: { phase: string; cancelled: boolean; aborted: boolean; preserved: boolean }[] = [];
  try {
    const first = await controller.render(sheet.sheetId, viewport), firstWindow = controller.current!, oldImages = [...firstWindow.element.querySelectorAll('img')];
    const firstNodes = firstWindow.nodes, stable = surface.view.innerHTML, frameStyle = surface.frame.style.cssText;
    for (const kind of ['identity', 'object']) {
      wrongRevision = kind === 'identity'; brokenObject = kind === 'object';
      try { await controller.render(sheet.sheetId, { ...viewport, heightPt: 320 }); errors.push('accepted'); }
      catch (error) { errors.push((error as { code: string }).code); }
      if (controller.current !== firstWindow || stable !== surface.view.innerHTML || surface.frame.style.cssText !== frameStyle) throw Error('A failed window changed the active view');
    }
    const pending = stall('cells'), obsolete = controller.render(sheet.sheetId, { ...viewport, scrollTopPt: 15 }); await pending.started;
    const retainedWhilePending = controller.current === firstWindow && surface.view.innerHTML === stable;
    const switched = await controller.render('controller-empty', { ...viewport, widthPt: 380, heightPt: 260 });
    const cancelled = await obsolete, switchedWindow = controller.current!, switchedHtml = surface.view.innerHTML;
    await pending.release(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const lateCellIgnored = controller.current === switchedWindow && surface.view.innerHTML === switchedHtml;
    const releasedOld = oldImages.every(image => !image.isConnected && !image.hasAttribute('src')) && !firstWindow.element.childNodes.length;
    const empty = { sheetId: switchedWindow.frame.sheetId, objects: switchedWindow.objects.length, width: parseFloat(surface.frame.style.width), height: parseFloat(surface.frame.style.height) };
    await controller.render(sheet.sheetId, viewport);
    const phase: Phase = snapshot.metadata.workbook.charts?.length ? 'charts' : 'assets';
    for (const fail of [false, true]) {
      const gate = stall(phase), work = controller.render(sheet.sheetId, { ...viewport, scrollTopPt: 10 }); await gate.started;
      await controller.render(sheet.sheetId, { ...viewport, scrollTopPt: 30 });
      const accepted = controller.current!, html = surface.view.innerHTML, cancelled = await work;
      await gate.release(fail); await new Promise<void>(resolve => setTimeout(resolve, 0));
      races.push({ phase, cancelled: !cancelled, aborted: gate.aborted(), preserved: accepted === controller.current && html === surface.view.innerHTML });
    }
    const secondSurface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
    const bounded = new WorkbookWindowController(secondSurface, { cells, assets, charts }, formatter, { maxNodes: firstNodes - 1 });
    let budget = '';
    try { await bounded.render(sheet.sheetId, viewport); budget = 'accepted'; } catch (error) { budget = (error as { code: string }).code; }
    const failedMountEmpty = secondSurface.view.childNodes.length === 0 && bounded.current === undefined; bounded.dispose(); secondSurface.dispose();
    const active = controller.current!, activeImages = [...active.element.querySelectorAll('img')], actual = surface.document.createTreeWalker(active.element, 0xffff_ffff);
    let actualNodes = 1; while (actual.nextNode()) actualNodes++;
    const nodes = active.nodes;
    const gate = stall(phase), last = controller.render(sheet.sheetId, viewport); await gate.started;
    controller.dispose(); controller.dispose(); const cancelledOnDispose = !await last;
    await gate.release(); await new Promise<void>(resolve => setTimeout(resolve, 0));
    const after = model.serialize();
    return { first, errors, retainedWhilePending, switched, obsoleteCancelled: !cancelled && pending.aborted(), empty, lateCellIgnored, releasedOld, races, budget, failedMountEmpty,
      nodes, actualNodes, cancelledOnDispose, released: [...created].every(url => revoked.has(url)), imagesCleared: activeImages.every(image => !image.hasAttribute('src')),
      emptyAfterDispose: !surface.view.childNodes.length && !surface.source.childNodes.length && controller.current === undefined,
      formatterStillOwned: formatter.format(1).kind === 'parts', unhandled,
      unchanged: [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])) };
  } finally {
    controller.dispose(); formatter.dispose(); surface.dispose(); container.remove();
    URL.createObjectURL = originalCreate; URL.revokeObjectURL = originalRevoke; window.removeEventListener('unhandledrejection', rejection);
  }
}
