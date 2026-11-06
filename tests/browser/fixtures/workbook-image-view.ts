import type { WorkbookSnapshot } from '@doffi/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { prepareWorkbookAssets } from '../../../packages/render/src/workbook/assets.js';
import { AxisDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { workbookViewport } from '../../../packages/render/src/workbook/viewport.js';
import { projectWorkbookDrawing } from '../../../packages/render/src/workbook/drawing.js';
import { createWorkbookDrawingFragment, workbookDrawingOrientation } from '../../../packages/render/src/workbook/drawing-view.js';
import { createWorkbookImageContent, decodeWorkbookImages } from '../../../packages/render/src/workbook/image-view.js';

let previous: { node: HTMLElement; dispose(): void } | undefined;
export function disposeImageView() { previous?.node.remove(); previous?.dispose(); previous = undefined; }

export async function measurePictureFills(snapshot: WorkbookSnapshot, resources: Readonly<Record<string, string>>, rtl: boolean) {
  disposeImageView();
  const scope = document.createElement('div'); document.body.append(scope);
  const source = new WorkbookContentSources(document, snapshot), objects = snapshot.metadata.workbook.objects!;
  const projection = source.project(scope, objects.map(object => object.nodeId)), paths = [...new Set(projection.assets.map(asset => asset.path))];
  const lease = await prepareWorkbookAssets({ snapshot, async readAsset(request) { return { ...request, bytes: Uint8Array.from(atob(resources[request.path]!), value => value.charCodeAt(0)) }; } }, paths, new AbortController().signal);
  const gallery = document.createElement('div'); gallery.id = 'workbook-picture-fill-probe'; gallery.style.cssText = 'position:absolute;left:0;top:0;display:flex;flex-wrap:wrap;width:1000px;gap:4px;background:white';
  const original = JSON.stringify(snapshot), images: HTMLImageElement[] = [], records = [];
  let accepted = false;
  try {
    for (const object of objects) {
      const element = projection.get(object.nodeId)!, orientation = workbookDrawingOrientation(element), rect = { x: 30, y: 55, width: 180, height: 90 };
      const path = projection.assets.find(asset => element.contains(projection.get(asset.nodeId)!))!.path;
      const view = createWorkbookImageContent(element, path, lease, { widthPt: rect.width, heightPt: rect.height, orientation }); images.push(view.image);
      const walker = document.createTreeWalker(view.element, 0xffff_ffff); let actualNodes = 1; while (walker.nextNode()) actualNodes++;
      const fragment = createWorkbookDrawingFragment(document, object.nodeId, { pane: 'body', rect, clip: { x: 0, y: 0, width: 240, height: 200 } }, 240, orientation, rtl, view.element);
      const cell = document.createElement('div'); cell.style.cssText = 'position:relative;width:240pt;height:200pt;overflow:hidden;background:white'; cell.append(fragment); gallery.append(cell);
      records.push({ cell, orientation, rotates: element.getAttribute('data-office-image-rotate-with-shape') !== 'false', nodes: view.nodes, actualNodes, image: view.image, fill: view.element.querySelector<HTMLElement>('[data-office-view=workbook-image-fill]') });
    }
    await decodeWorkbookImages(images, new AbortController().signal); document.body.append(gallery);
    const results = records.map(record => {
      const box = record.cell.getBoundingClientRect(), angle = record.orientation.rotationDegrees * Math.PI / 180, cosine = Math.cos(angle), sine = Math.sin(angle), sx = record.orientation.flipH ? -1 : 1, sy = record.orientation.flipV ? -1 : 1;
      const width = Math.abs(cosine)*180 + Math.abs(sine)*90, height = Math.abs(sine)*180 + Math.abs(cosine)*90;
      // Sample on each side of the original bitmap's quadrant boundaries.
      // Page-oriented fill samples stay in page axes, even after reflection.
      const samples = [[-.15, -.15], [.15, -.15], [-.15, .15], [.15, .15]].map(([u, v]) => {
        const x = record.rotates ? cosine*u!*180*sx - sine*v!*90*sy : u!*width*sx;
        const y = record.rotates ? sine*u!*180*sx + cosine*v!*90*sy : v!*height*sy;
        return { x: box.x + (120 + x)/.75, y: box.y + (100 + y)/.75 };
      });
      const body = record.cell.querySelector<HTMLElement>('[data-office-view=workbook-drawing-body]')!, bodyMatrix = new DOMMatrix(getComputedStyle(body).transform);
      const fillMatrix = record.fill ? new DOMMatrix(getComputedStyle(record.fill).transform) : new DOMMatrix(), matrix = bodyMatrix.multiply(fillMatrix);
      return { rotates: record.rotates, orientation: record.orientation, nodes: record.nodes, actualNodes: record.actualNodes, samples, matrix: [matrix.a, matrix.b, matrix.c, matrix.d] };
    });
    const errors: string[] = [], first = projection.get(objects[0]!.nodeId)!;
    for (const value of ['yes', '', '0']) {
      const bad = first.cloneNode(true) as HTMLElement; bad.setAttribute('data-office-image-rotate-with-shape', value);
      try { createWorkbookImageContent(bad, paths[0]!, lease); } catch (error) { errors.push((error as { code: string }).code); }
    }
    const bad = first.cloneNode(true) as HTMLElement; bad.setAttribute('data-office-image-rotate-with-shape', 'false');
    for (const frame of [undefined, { widthPt: -1, heightPt: 90, orientation: workbookDrawingOrientation(first) }, { widthPt: 1e-300, heightPt: 1e300, orientation: { ...workbookDrawingOrientation(first), rotationDegrees: 45 } }]) {
      try { createWorkbookImageContent(bad, paths[0]!, lease, frame); } catch (error) { errors.push((error as { code: string }).code); }
    }
    const degenerate = createWorkbookImageContent(bad, paths[0]!, lease, { widthPt: 0, heightPt: 90, orientation: workbookDrawingOrientation(first) }); degenerate.image.removeAttribute('src');
    previous = { node: gallery, dispose: () => { images.forEach(image => image.removeAttribute('src')); lease.dispose(); } }; accepted = true;
    return { results, errors, unchanged: original === JSON.stringify(snapshot), degenerate: [degenerate.scaleX, degenerate.scaleY] };
  } finally { source.dispose(); scope.remove(); if (!accepted) { gallery.remove(); lease.dispose(); } }
}

export async function measureImageView(snapshot: WorkbookSnapshot, resources: Readonly<Record<string, string>>, rtl: boolean) {
  disposeImageView();
  const scope = document.createElement('div'); document.body.append(scope);
  const object = snapshot.metadata.workbook.objects!.find(object => object.kind === 'image')!, source = new WorkbookContentSources(document, snapshot), projection = source.project(scope, [object.nodeId]);
  const paths = projection.assets.map(asset => asset.path), signal = new AbortController().signal;
  const lease = await prepareWorkbookAssets({ snapshot, async readAsset(request) { return { ...request, bytes: Uint8Array.from(atob(resources[request.path]!), value => value.charCodeAt(0)) }; } }, paths, signal);
  const candidate = document.createElement('div'); candidate.id = 'workbook-image-probe'; candidate.style.cssText = 'position:relative;width:240pt;height:180pt;background:#eee;overflow:hidden';
  let accepted = false;
  try {
    const element = projection.get(object.nodeId)!, orientation = { ...workbookDrawingOrientation(element), flipH: true }, rect = { x: 10, y: 5, width: 200, height: 150 };
    const dimensions = { rows: new AxisDimensions(100, 20, []), columns: new AxisDimensions(100, 40, []) };
    const layout = workbookViewport(dimensions, { rows: 2, columns: 1 }, { widthPt: 240, heightPt: 180, scrollLeftPt: 20, scrollTopPt: 10 }, 200);
    const parts = projectWorkbookDrawing(rect, orientation.rotationDegrees, layout), images: HTMLImageElement[] = [];
    for (const part of parts) {
      const view = createWorkbookImageContent(element, paths[0]!, lease); images.push(view.image);
      const fragment = createWorkbookDrawingFragment(document, object.nodeId, part, layout.widthPt, orientation, rtl, view.element); candidate.append(fragment);
    }
    await decodeWorkbookImages(images, signal); scope.append(candidate);
    const panes = [...candidate.children].map(fragment => {
      const outer = fragment as HTMLElement, body = outer.firstElementChild as HTMLElement, image = body.querySelector('img')!;
      const style = getComputedStyle(body), box = body.getBoundingClientRect(), matrix = new DOMMatrix(style.transform);
      return { pane: outer.dataset.pane, clip: [outer.offsetLeft, outer.offsetTop, outer.offsetWidth, outer.offsetHeight].map(value => value * .75), body: [parseFloat(body.style.left), parseFloat(body.style.top), parseFloat(body.style.width), parseFloat(body.style.height)],
        transformed: [box.width * .75, box.height * .75], matrix: [matrix.a, matrix.b, matrix.c, matrix.d],
        crop: [image.style.left, image.style.top, image.style.width, image.style.height].map(parseFloat), natural: [image.naturalWidth, image.naturalHeight], nodeId: image.dataset.sourceNodeId };
    });
    // Failing/aborted candidates cannot change the accepted image subtree.
    const bad = document.createElement('img'); bad.src = 'data:image/png;base64,AA==';
    let invalid = '', cancelled = '', connected = '';
    try { await decodeWorkbookImages([bad], signal); } catch (error) { invalid = (error as { code: string }).code; }
    const waiting = document.createElement('img'); waiting.src = lease.url(paths[0]!)!; let finish!: () => void;
    waiting.decode = () => new Promise<void>(resolve => { finish = resolve; });
    const controller = new AbortController(), work = decodeWorkbookImages([waiting], controller.signal);
    controller.abort(); try { await work; } catch (error) { cancelled = (error as { code: string }).code; } finish(); await Promise.resolve();
    try { await decodeWorkbookImages(images, signal); } catch (error) { connected = (error as { code: string }).code; }
    const kept = images.every(image => image.isConnected && image.hasAttribute('src'));
    const sourceCrops = ['left', 'top', 'right', 'bottom'].map(side => Number(element.getAttribute('data-office-crop-' + side) ?? 0));
    const badGeometry: string[] = [];
    for (const [name, value] of [['data-office-crop-left', '1'], ['data-office-crop-left', 'NaN'], ['data-office-crop-left', '.99999999']]) {
      const changed = element.cloneNode(true) as HTMLElement; changed.removeAttribute('data-office-crop-right'); changed.setAttribute(name!, value!);
      try { createWorkbookImageContent(changed, paths[0]!, lease); } catch (error) { badGeometry.push((error as { code: string }).code); }
    }
    const extreme = element.cloneNode(true) as HTMLElement; extreme.setAttribute('data-office-crop-left', '.99999'); extreme.removeAttribute('data-office-crop-right');
    const nativeCrop = createWorkbookImageContent(extreme, paths[0]!, lease); nativeCrop.image.removeAttribute('src');
    const extremeAccepted = parseFloat(nativeCrop.image.style.width) > 9_000_000;
    let active = 0, peak = 0, decoded = 0;
    const many = Array.from({ length: 17 }, () => {
      const image = document.createElement('img'); image.decode = () => new Promise<void>(resolve => { active++; peak = Math.max(peak, active); setTimeout(() => { active--; decoded++; resolve(); }, 0); }); return image;
    });
    await decodeWorkbookImages(many, signal);
    // Keep only the accepted visual for the test screenshot; its lease is
    // released when the next fixture replaces it or the test page is closed.
    document.body.append(candidate); previous = { node: candidate, dispose: lease.dispose }; accepted = true;
    return { panes, sourceCrops, sourceRotation: orientation.rotationDegrees, invalid, cancelled, connected, kept, cleared: !bad.hasAttribute('src') && !waiting.hasAttribute('src'), badGeometry, extremeAccepted, peak, decoded };
  } finally { source.dispose(); scope.remove(); if (!accepted) lease.dispose(); }
}
