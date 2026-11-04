import type { WorkbookSnapshot } from '@dagent-office/core';
import { WorkbookContentSources } from '../../../packages/render/src/workbook/content-source.js';
import { prepareWorkbookAssets } from '../../../packages/render/src/workbook/assets.js';
import { AxisDimensions } from '../../../packages/render/src/workbook/dimensions.js';
import { workbookViewport } from '../../../packages/render/src/workbook/viewport.js';
import { projectWorkbookDrawing } from '../../../packages/render/src/workbook/drawing.js';
import { createWorkbookDrawingFragment, workbookDrawingOrientation } from '../../../packages/render/src/workbook/drawing-view.js';
import { createWorkbookImageContent, decodeWorkbookImages } from '../../../packages/render/src/workbook/image-view.js';

let previous: { node: HTMLElement; dispose(): void } | undefined;
export function disposeImageView() { previous?.node.remove(); previous?.dispose(); previous = undefined; }

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
