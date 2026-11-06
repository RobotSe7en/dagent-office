import type { WorkbookSnapshot } from '@dagent-office/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookDrawingText } from '../../../packages/render/src/workbook/drawing-text.js';
import { layoutWorkbookDrawingTexts } from '../../../packages/render/src/workbook/drawing-text-layout.js';
import { createWorkbookObjectGeometry } from '../../../packages/render/src/workbook/object-geometry.js';
import { readWorkbookViewRange, selectWorkbookView } from '../../../packages/render/src/workbook/selection.js';

let releasePixels: (() => void) | undefined;
export function disposeWorkbookSurfacePixels(): void { releasePixels?.(); releasePixels = undefined; }
export async function mountWorkbookSurfacePixels(snapshot: WorkbookSnapshot) {
  disposeWorkbookSurfacePixels();
  const container = document.createElement('div'); container.id = 'workbook-surface-pixel-probe';
  container.style.cssText = 'position:fixed;left:0;top:0;z-index:100000;width:400px;height:200px;background:white'; document.body.append(container);
  let surface: Awaited<ReturnType<typeof createWorkbookSurface>> | undefined, sources: WorkbookContentSources | undefined;
  try {
    surface = await createWorkbookSurface(container, snapshot, { widthPt: 300, heightPt: 150 });
    sources = new WorkbookContentSources(surface.document, snapshot);
    const object = snapshot.metadata.workbook.objects!.find(object => object.kind === 'geometry')!, projection = sources.project(surface.source, [object.nodeId]), source = projection.get(object.nodeId)!;
    for (const [name, value] of Object.entries({ kind: 'linear', angle: '0', scaled: 'false', 'stop-count': '2', 'stop-0-position': '0', 'stop-0-color': '#ff0000', 'stop-1-position': '1', 'stop-1-color': '#0000ff' })) source.style.setProperty('--office-fill-' + name, value);
    source.style.setProperty('--office-line-kind', 'none');
    const geometry = createWorkbookObjectGeometry(source, object, 300, 150, 'isolated-pixel-gradient'); surface.view.append(geometry.element);
    surface.frame.style.setProperty('visibility', 'visible', 'important');
    releasePixels = () => { sources!.dispose(); surface!.dispose(); container.remove(); };
  } catch (error) { sources?.dispose(); surface?.dispose(); container.remove(); throw error; }
}

export async function measureWorkbookSurface(snapshot: WorkbookSnapshot) {
  const container = document.createElement('div'), hostStyles = document.createElement('style');
  hostStyles.textContent = 'p{font-size:99px!important;color:red!important}iframe{width:1px!important;height:1px!important;border:9px solid red!important}';
  container.style.cssText = 'position:relative;width:400px;height:200px'; document.head.append(hostStyles); document.body.append(container);
  const options = { widthPt: 300, heightPt: 150 }, surfaces: Awaited<ReturnType<typeof createWorkbookSurface>>[] = [], values = [], errors: string[] = [];
  try {
    for (const color of ['rgb(13, 75, 91)', 'rgb(9, 33, 77)']) {
      const css = `:root{--surface-color:${color};font-family:"Liberation Sans";font-size:16px}body{line-height:1.25;width:777px!important;height:888px!important;margin:40px!important;padding:30px!important}p{width:999pt!important;font-size:14pt!important;color:var(--surface-color)!important;line-height:normal!important}div{width:888pt!important;height:444pt!important;border:8pt solid red!important}svg{width:777pt!important;height:666pt!important}office-workbook-view-surface{width:9999pt!important;height:8888pt!important;font-size:99px!important}`;
      const input = { ...snapshot, styles: [...snapshot.styles, { path: 'styles/surface-probe.css', css }] }, controller = new AbortController();
      const surface = await createWorkbookSurface(container, input, { ...options, signal: controller.signal }); surfaces.push(surface);
      const sources = new WorkbookContentSources(surface.document, input), object = input.metadata.workbook.objects!.find(object => object.kind === 'geometry')!;
      const projection = sources.project(surface.source, [object.nodeId]), source = projection.get(object.nodeId)!, before = source.outerHTML;
      try {
        const text = createWorkbookDrawingText(source, object), result = layoutWorkbookDrawingTexts(surface.view, [{ source, object, text, widthPt: 120, heightPt: 90 }]).objects[0]!;
        const geometry = createWorkbookObjectGeometry(source, object, 120, 90, 'surface-geometry'); result.element.append(geometry.element); surface.view.append(result.element);
        const view = surface.document.defaultView!, sourceStyle = view.getComputedStyle(source.querySelector('p')!), paragraph = text.element.querySelector('p')!;
        const hidden = getComputedStyle(surface.frame).visibility;
        surface.frame.style.setProperty('visibility', 'visible', 'important');
        const selected = text.projections[0]!.node, selection = view.getSelection()!;
        // Backward selections retain the same exact ordered source range.
        selectWorkbookView(surface.view, { node: selected, offset: selected.length }, { node: selected, offset: 0 });
        const range = readWorkbookViewRange(surface.view)!, selectionErrors: string[] = [];
        for (const focus of [{ node: selected, offset: selected.length + 1 }, { node: source, offset: 0 }]) {
          try { selectWorkbookView(surface.view, { node: selected, offset: 0 }, focus); selectionErrors.push('accepted'); } catch (e) { selectionErrors.push((e as { code: string }).code); }
        }
        const sameSelection = readWorkbookViewRange(surface.view)?.toString() === range.toString();
        controller.abort();
        values.push({ sourceColor: sourceStyle.color, sourceFont: sourceStyle.fontSize, sourceWidth: sourceStyle.width,
          frame: [surface.frame.clientWidth, surface.frame.clientHeight], view: [surface.view.clientWidth, surface.view.clientHeight],
          object: [result.element.clientWidth, result.element.clientHeight], geometry: [geometry.element.clientWidth, geometry.element.clientHeight],
          paragraphFont: view.getComputedStyle(paragraph).fontSize, paragraphWidth: paragraph.clientWidth, border: view.getComputedStyle(result.element).borderLeftWidth,
          hidden, connectedAfterReadyAbort: surface.frame.isConnected, unchanged: before === source.outerHTML,
          selection: { text: selection.toString(), expected: selected.data, range: range.startContainer === selected && range.endContainer === selected && range.startOffset === 0 && range.endOffset === selected.length, errors: selectionErrors, unchanged: sameSelection } });
        const outside = surface.document.createElement('span'), outsideText = surface.document.createTextNode('Outside'); outside.append(outsideText); outside.style.cssText = 'position:absolute;left:0;top:0'; surface.document.body.append(outside);
        selection.setBaseAndExtent(outsideText, 0, outsideText, outsideText.length);
        if (readWorkbookViewRange(surface.view) !== undefined) throw Error('Another surface selection leaked into the display projection');
        outside.remove();
      } finally { sources.dispose(); }
    }
    for (const kind of ['budget', 'invalid', 'already-cancelled', 'cancelled']) {
      const controller = new AbortController(); if (kind === 'already-cancelled') controller.abort();
      try {
        const pending = createWorkbookSurface(container, snapshot, { ...options, ...(kind === 'budget' ? { maxCssCodeUnits: 1 } : {}), ...(kind === 'invalid' ? { widthPt: NaN } : {}), signal: controller.signal });
        if (kind === 'cancelled') controller.abort();
        (await pending).dispose(); errors.push('accepted');
      } catch (error) { errors.push((error as { code: string }).code); }
    }
    const kept = surfaces.every(surface => surface.frame.isConnected) && container.querySelectorAll('iframe').length === 2;
    // Disposing one workbook cannot change the second document's styles/view.
    surfaces[0]!.dispose(); surfaces[0]!.dispose();
    const remaining = surfaces[1]!, independent = remaining.frame.isConnected && remaining.view.querySelector('p') !== null && container.querySelectorAll('iframe').length === 1;
    remaining.dispose();
    return { values, errors, kept, independent, released: container.childElementCount === 0 && surfaces.every(surface => !surface.view.childNodes.length && !surface.source.childNodes.length && !surface.document.head.childNodes.length && !surface.document.body.childNodes.length) };
  } finally { for (const surface of surfaces) surface.dispose(); container.remove(); hostStyles.remove(); }
}
