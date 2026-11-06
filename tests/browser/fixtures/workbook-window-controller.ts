import { decodeTestPackage, type TestPackage } from './package-transport.js';
import { createWorkbookNumberFormatter, openDocument } from '@doffi/core';
import type { WorkbookDocumentMetadata } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { modelWindowSource } from '../../../packages/render/src/workbook/window-source.js';
import { modelAssetSource } from '../../../packages/render/src/workbook/assets.js';
import { modelChartSource } from '../../../packages/render/src/workbook/chart-source.js';
import { WorkbookWindowController } from '../../../packages/render/src/workbook/window-controller.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookCommentText } from '../../../packages/render/src/workbook/comment-text.js';
import { resolvePreviewFontsInScope } from '../../../packages/render/src/fonts.js';
import { readWorkbookViewRange, selectWorkbookView } from '../../../packages/render/src/workbook/selection.js';

export async function windowControllerMeasurement(entries: TestPackage) {
  const files = decodeTestPackage(entries), decoder = new TextDecoder(), encoder = new TextEncoder();
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

export async function annotationWindowMeasurement(entries: TestPackage) {
  const files = decodeTestPackage(entries);
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize(), sheet = snapshot.metadata.workbook.sheets[0]!;
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 });
  const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
  const controller = new WorkbookWindowController(surface, { cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model) }, formatter);
  try {
    const viewport = { widthPt: 420, heightPt: 300, scrollLeftPt: 0, scrollTopPt: 0 };
    await controller.render(sheet.sheetId, viewport);
    const current = controller.current!, visible = current.annotations.visible.map(item => ({ id: item.annotation.annotationId, status: item.annotation.status,
      format: item.native?.format, ranges: item.regions.map(region => region.rangeIndex),
      body: item.annotation.bodyNodeId ? model.getNode(item.annotation.bodyNodeId)!.text : item.annotation.instruction,
      replies: item.annotation.replies?.map(reply => reply.bodyNodeId ? model.getNode(reply.bodyNodeId)!.text : reply.instruction) ?? [] }));
    const expected = snapshot.annotations.filter(annotation => annotation.status !== 'needs_reanchor').map(annotation => annotation.annotationId);
    const recovery = current.annotations.recovery.map(item => item.annotation.annotationId);
    await controller.render(sheet.sheetId, { ...viewport, scrollLeftPt: current.frame.layout.maxScrollLeftPt, scrollTopPt: current.frame.layout.maxScrollTopPt });
    const offscreen = controller.current!.annotations.visible.length, after = model.serialize();
    return { visible, expected, recovery, expectedRecovery: snapshot.annotations.filter(annotation => annotation.status === 'needs_reanchor').map(annotation => annotation.annotationId), offscreen,
      unchanged: [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])) };
  } finally { controller.dispose(); formatter.dispose(); surface.dispose(); container.remove(); }
}

export async function commentTextMeasurement(entries: TestPackage, rich: boolean | 'singleAccounting' | 'doubleAccounting' = false) {
  const files = decodeTestPackage(entries), encoder = new TextEncoder(), decoder = new TextDecoder();
  if (rich) {
    const manifest = JSON.parse(decoder.decode(files.get('manifest.json')!)) as { content: string; styles: string[] };
    const template = document.createElement('template'); template.innerHTML = decoder.decode(files.get(manifest.content)!);
    const story = template.content.querySelector('[data-office-kind="comment"]')!;
    story.append(document.createTextNode(' \n'));
    const paragraph = document.createElement('p'); paragraph.dataset.nodeId = 'review-comment-paragraph'; paragraph.style.whiteSpace = 'pre-wrap';
    const span = document.createElement('span'); span.dataset.nodeId = 'review-comment-run';
    span.style.cssText = '--office-font-underline:' + (typeof rich === 'string' ? rich : 'double') + ';--office-font-strike:true;--office-cell-foreground-color:#ff0000;--office-cell-foreground-tint:0;font-size:14pt';
    span.textContent = '混排😀禰󠄀';
    const line = document.createElement('br'); line.dataset.nodeId = 'review-comment-break';
    const tail = document.createElement('span'); tail.dataset.nodeId = 'review-comment-tail'; tail.style.cssText = '--office-font-underline:none;--office-font-strike:false'; tail.textContent = 'No underline';
    paragraph.append(span, line, tail); story.append(paragraph);
    files.set(manifest.content, encoder.encode('<!doctype html><html><head><meta charset="utf-8"><link rel="stylesheet" href="' + manifest.styles[0]! + '"></head><body>' + template.content.querySelector('[data-office-kind="workbook-content"]')!.outerHTML + '</body></html>'));
  }
  const model = openDocument(files, { kind: 'workbook' }), snapshot = model.snapshot(), initial = model.serialize();
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 420, heightPt: 300 }), sources = new WorkbookContentSources(surface.document, snapshot);
  const ids = snapshot.metadata.workbook.stories!.map(story => story.rootNodeId), projection = sources.project(surface.source, ids);
  try {
    const results = [];
    for (const id of ids) {
      const source = projection.get(id)!, before = source.outerHTML, rendered = createWorkbookCommentText(source);
      const detached = !rendered.element.isConnected; surface.view.append(rendered.element);
      const fonts = resolvePreviewFontsInScope(surface.document, rendered.element, '[data-office-text-style],[data-office-script]');
      const expected = model.getNode(id)!.text, walker = surface.document.createTreeWalker(rendered.element, 0xffff_ffff); let count = 1; while (walker.nextNode()) count++;
      const offsets = rendered.projections.map(piece => ({ from: piece.from, to: piece.to, text: piece.node.data, expected: expected.slice(piece.from, piece.to), sourceNodeId: piece.sourceNodeId }));
      const pieces = rendered.projections, first = pieces[0]!, last = pieces.at(-1)!;
      selectWorkbookView(surface.view, { node: first.node, offset: 0 }, { node: last.node, offset: last.node.length });
      const selected = readWorkbookViewRange(surface.view);
      const formats = [...rendered.element.querySelectorAll<HTMLElement>('[data-office-script]')].map(node => ({ text: node.textContent, from: rendered.projections.find(piece => node.contains(piece.node))!.from, underline: node.style.textDecorationLine, style: node.style.textDecorationStyle,
        strike: node.firstElementChild instanceof surface.document.defaultView!.HTMLElement ? (node.firstElementChild as HTMLElement).style.textDecorationLine : '', color: surface.document.defaultView!.getComputedStyle(node).color }));
      const errors = [];
      for (const [nodes, text] of [[rendered.nodes - 1, 8 * 1024 * 1024], [50_000, 1]]) {
        try { createWorkbookCommentText(source, nodes, text); errors.push('accepted'); } catch (error) { errors.push((error as { code: string }).code); }
      }
      results.push({ id, detached, unchangedSource: source.outerHTML === before, text: rendered.element.textContent, expected, offsets, nodes: rendered.nodes, actualNodes: count, formats, fonts: fonts.length, errors,
        canonicalIds: rendered.element.querySelectorAll('[data-node-id]').length, selected: selected?.startContainer === first.node && selected.endContainer === last.node });
      surface.document.defaultView!.getSelection()!.removeAllRanges(); rendered.element.remove();
    }
    const after = model.serialize();
    return { results, unchanged: [...initial].every(([path, bytes]) => bytes.length === after.get(path)!.length && bytes.every((byte, index) => byte === after.get(path)![index])) };
  } finally { projection.dispose(); sources.dispose(); surface.dispose(); container.remove(); }
}
