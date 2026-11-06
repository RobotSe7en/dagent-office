import type { WorkbookSnapshot } from '@doffi/core';
import { createWorkbookSurface } from '../../../packages/render/src/workbook/surface.js';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookObjectGeometry } from '../../../packages/render/src/workbook/object-geometry.js';
import { workbookDrawingShadow } from '../../../packages/render/src/workbook/drawing-shadow.js';
import { createWorkbookDrawingFragment, workbookDrawingOrientation } from '../../../packages/render/src/workbook/drawing-view.js';

export async function measureWorkbookShadows(snapshot: WorkbookSnapshot) {
  const container = document.createElement('div'); document.body.append(container);
  const surface = await createWorkbookSurface(container, snapshot, { widthPt: 300, heightPt: 180 }), sources = new WorkbookContentSources(surface.document, snapshot);
  const objects = snapshot.metadata.workbook.objects!, projection = sources.project(surface.source, objects.map(object => object.nodeId));
  try {
    const cases = objects.map(object => {
      const source = projection.get(object.nodeId)!, before = source.outerHTML, orientation = workbookDrawingOrientation(source);
      const shadow = workbookDrawingShadow(source, orientation, 120, 60, { pixelsPerPoint: 1 / .75 });
      const content = surface.document.createElement('div'); content.style.filter = shadow.filter;
      const fragment = createWorkbookDrawingFragment(surface.document, object.nodeId, { pane: 'body', rect: { x: 30, y: 30, width: 120, height: 60 }, clip: { x: 0, y: 0, width: 300, height: 180 } }, 300, orientation, false, content);
      surface.view.append(fragment); const body = fragment.querySelector<HTMLElement>('[data-office-view="workbook-drawing-body"]')!;
      const matrix = new DOMMatrix(surface.document.defaultView!.getComputedStyle(body).transform), offset = shadow.localOffsetPt;
      const result = { worldOffsetPt: [matrix.a * offset.x + matrix.c * offset.y, matrix.b * offset.x + matrix.d * offset.y], outset: shadow.outsetPt, pixels: shadow.rasterPixels, approximate: shadow.approximateBlur,
        filterEnabled: surface.document.defaultView!.getComputedStyle(content).filter !== 'none', unchanged: before === source.outerHTML };
      const exact = workbookDrawingShadow(source, orientation, 120, 60, { pixelsPerPoint: 1 / .75, maxRasterPixels: shadow.rasterPixels });
      if (exact.rasterPixels !== shadow.rasterPixels) throw Error('Shadow budget boundary changed'); fragment.remove(); return result;
    });
    const source = projection.get(objects[0]!.nodeId)!, orientation = workbookDrawingOrientation(source), before = source.getAttribute('style'), errors: string[] = [];
    const restore = () => { if (before === null) source.removeAttribute('style'); else source.setAttribute('style', before); };
    for (const [name, value] of [['blur', '-2pt'], ['blur', '2px'], ['direction', '5deg'], ['rotate-with-shape', 'yes'], ['align', 'center'], ['distance', '1e309pt']]) {
      try { source.style.setProperty('--office-shadow-' + name, value!); workbookDrawingShadow(source, orientation, 120, 60); errors.push('accepted'); } catch (e) { errors.push((e as { code: string }).code); } finally { restore(); }
    }
    const expectedPixels = cases[0]!.pixels;
    try { workbookDrawingShadow(source, orientation, 120, 60, { pixelsPerPoint: 1 / .75, maxRasterPixels: expectedPixels - 1 }); errors.push('accepted'); } catch (e) { errors.push((e as { code: string }).code); }
    try { workbookDrawingShadow(source, orientation, 1_000_000, 1_000_000); errors.push('accepted'); } catch (e) { errors.push((e as { code: string }).code); }
    try { source.style.setProperty('--office-shadow-kind', 'inner'); workbookDrawingShadow(source, orientation, 120, 60); errors.push('accepted'); } catch (e) { errors.push((e as { code: string }).code); } finally { restore(); }
    source.style.setProperty('--office-shadow-kind', 'none');
    const none = workbookDrawingShadow(source, orientation, 1_000_000, 1_000_000, { maxRasterPixels: 1 }); restore();
    return { cases, errors, none, restored: source.getAttribute('style') === before };
  } finally { sources.dispose(); surface.dispose(); container.remove(); }
}

let release: (() => void) | undefined;
export function disposeWorkbookShadowPixels(): void { release?.(); release = undefined; }
export async function mountWorkbookShadowPixels(snapshot: WorkbookSnapshot, index: number) {
  disposeWorkbookShadowPixels();
  const container = document.createElement('div'); container.id = 'workbook-shadow-pixel-probe'; container.style.cssText = 'position:fixed;left:0;top:0;z-index:100000;width:480px;height:240px;background:white'; document.body.append(container);
  let surface: Awaited<ReturnType<typeof createWorkbookSurface>> | undefined, sources: WorkbookContentSources | undefined;
  try {
    surface = await createWorkbookSurface(container, snapshot, { widthPt: 360, heightPt: 180 }); sources = new WorkbookContentSources(surface.document, snapshot);
    const object = snapshot.metadata.workbook.objects![index]!, projection = sources.project(surface.source, [object.nodeId]), source = projection.get(object.nodeId)!, orientation = workbookDrawingOrientation(source);
    const shadow = workbookDrawingShadow(source, orientation, 120, 60), content = surface.document.createElement('div');
    content.style.cssText = `position:relative;width:120pt;height:60pt;filter:${shadow.filter}`;
    content.append(createWorkbookObjectGeometry(source, object, 120, 60, 'native-shadow-blur').element);
    surface.view.append(createWorkbookDrawingFragment(surface.document, object.nodeId, { pane: 'body', rect: { x: 40, y: 40, width: 120, height: 60 }, clip: { x: 0, y: 0, width: 360, height: 180 } }, 360, orientation, false, content));
    surface.frame.style.setProperty('visibility', 'visible', 'important'); release = () => { sources!.dispose(); surface!.dispose(); container.remove(); };
    return { approximate: shadow.approximateBlur, outset: shadow.outsetPt, pixels: shadow.rasterPixels };
  } catch (error) { sources?.dispose(); surface?.dispose(); container.remove(); throw error; }
}
