import { createAssetDescriptor, createWorkbook, openDocument } from '@dagent-office/core';
import type { WorkbookManifest } from '@dagent-office/core';
import { modelAssetSource, prepareWorkbookAssets } from '../../../packages/render/src/workbook/assets.js';

export async function assetMeasurement(scope: HTMLElement, base64: string, mimeType: string) {
  const bytes = Uint8Array.from(atob(base64), character => character.charCodeAt(0)), files = new Map(createWorkbook({ documentId: 'asset-preview' }));
  const manifest = JSON.parse(new TextDecoder().decode(files.get('manifest.json')!)) as WorkbookManifest, path = `assets/sample.${mimeType.split('/')[1]}`;
  manifest.assets = [createAssetDescriptor(path, mimeType, bytes)]; files.set(path, bytes); files.set('manifest.json', new TextEncoder().encode(JSON.stringify(manifest)));
  const model = openDocument(files, { kind: 'workbook' }), source = modelAssetSource({ ...model, serialize() { throw Error('Asset reads must not serialize cell chunks'); } });
  const first = await prepareWorkbookAssets(source, [path, path], new AbortController().signal), second = await prepareWorkbookAssets(source, [path], new AbortController().signal);
  const images: HTMLImageElement[] = [];
  try {
    const firstUrl = first.url(path)!, secondUrl = second.url(path)!;
    for (const url of [firstUrl, secondUrl]) { const image = document.createElement('img'); image.src = url; images.push(image); scope.append(image); await image.decode(); }
    const dimensions = images.map(image => [image.naturalWidth, image.naturalHeight]);
    first.dispose(); first.dispose(); let revoked = false;
    try { await fetch(firstUrl); } catch { revoked = true; }
    const secondBytes = new Uint8Array(await (await fetch(secondUrl)).arrayBuffer());
    const isolated = secondBytes.length === bytes.length && secondBytes.every((byte, index) => byte === bytes[index]);
    const cancelled = new AbortController(); cancelled.abort(); let failure = '';
    try { await prepareWorkbookAssets(source, [path], cancelled.signal); } catch (error) { failure = (error as { code: string }).code; }
    return { count: first.count, byteLength: first.byteLength, pixels: first.pixels, distinct: firstUrl !== secondUrl, dimensions, revoked, isolated, failure };
  } finally { for (const image of images) image.remove(); first.dispose(); second.dispose(); }
}
