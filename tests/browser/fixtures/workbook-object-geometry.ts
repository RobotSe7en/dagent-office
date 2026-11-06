import type { WorkbookSnapshot } from '@doffi/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { createWorkbookObjectGeometry } from '../../../packages/render/src/workbook/object-geometry.js';
import { createDrawingGeometryView } from '../../../packages/render/src/drawing-geometry.js';

export function measureObjectGeometry(snapshot: WorkbookSnapshot) {
  const scope = document.createElement('div'), style = document.createElement('style'); style.textContent = snapshot.styles.map(style => style.css).join('\n');
  document.head.append(style); document.body.append(scope);
  const sources = new WorkbookContentSources(document, snapshot), objects = snapshot.metadata.workbook.objects!.filter(object => object.kind !== 'chart'), projection = sources.project(scope, objects.map(object => object.nodeId));
  try {
    const results = objects.map((object, index) => {
      const source = projection.get(object.nodeId)!, { element, geometry } = createWorkbookObjectGeometry(source, object, 200, 100, `workbook-geometry-${index}`);
      return { kind: object.kind, preset: element.getAttribute('data-sdk-preset'), paths: [...element.querySelectorAll(':scope > path')].map(path => ({ d: path.getAttribute('d'), fill: path.getAttribute('fill'), stroke: path.getAttribute('stroke'), width: path.getAttribute('stroke-width') })), textRect: geometry.textRect };
    });
    const source = projection.get(objects[1]!.nodeId)!, computed = getComputedStyle(source);
    const geometry = createWorkbookObjectGeometry(source, objects[1]!, 200, 100, 'geometry-budget').geometry;
    const errors: string[] = [];
    for (const maxNodes of [0, 1]) try { createDrawingGeometryView(document, computed, { geometry, preset: 'roundRect', id: 'limited', widthPt: 200, heightPt: 100, maxNodes }); } catch (error) { errors.push((error as { code: string }).code); }
    // Exercise shared gradient/arrow allocation with two paths. Marker targets
    // must be unique even when a native preset has more than one contour.
    const paint = document.createElement('div'); paint.style.cssText = '--office-fill-kind:linear;--office-fill-stop-count:2;--office-fill-stop-0-position:0;--office-fill-stop-0-color:#ff0000;--office-fill-stop-1-position:1;--office-fill-stop-1-color:#0000ff;--office-line-color:#00ff00;--office-line-start-arrow:triangle;--office-line-end-arrow:oval'; scope.append(paint);
    const multi = { ...geometry, paths: [...geometry.paths, ...geometry.paths] }, input = { geometry: multi, preset: 'test', id: 'multi-path', widthPt: 200, heightPt: 100, maxNodes: 19 };
    const svg = createDrawingGeometryView(document, getComputedStyle(paint), input), ids = [...svg.querySelectorAll('[id]')].map(node => node.id);
    try { createDrawingGeometryView(document, getComputedStyle(paint), { ...input, maxNodes: 18 }); } catch (error) { errors.push((error as { code: string }).code); }
    const arrows = [...svg.querySelectorAll(':scope > path')].flatMap(path => ['marker-start', 'marker-end'].map(name => path.getAttribute(name)));
    return { results, errors, nodes: svg.querySelectorAll('*').length + 1, ids, arrows, stops: [...svg.querySelectorAll('stop')].map(stop => stop.getAttribute('stop-color')) };
  } finally { sources.dispose(); scope.remove(); style.remove(); }
}
