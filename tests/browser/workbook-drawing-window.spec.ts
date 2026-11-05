import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type { drawingWindowMeasurement } from './fixtures/workbook-drawing-window.js';

declare global { interface Window { workbookDrawingWindow: { drawingWindowMeasurement: typeof drawingWindowMeasurement } } }
let code = '';
const inputs: { name: string; files: [string, number[]][] }[] = [];
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-drawing-window.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingWindow' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing-window bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`), { documentId: 'drawing-window' });
    inputs.push({ name: prefix + name, files: [...files].map(([path, bytes]) => [path, [...bytes]]) });
  }
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: code }); });
for (const rtl of [false, true]) test(`worksheet objects compose above cells across four ${rtl ? 'RTL' : 'LTR'} panes with isolated candidate resources`, async ({ page }) => {
  for (const input of inputs) {
    const result = await page.evaluate(({ files, rtl }) => window.workbookDrawingWindow.drawingWindowMeasurement(files, rtl), { files: input.files, rtl });
    expect(result, input.name).toMatchObject({ detached: true, unchanged: true, preserved: true, cleaned: true, released: true, clearedImages: true, removed: true, uniqueIds: true, cancelled: 'OperationCancelled', unhandled: [] });
    expect(result.errors).toEqual(Array(6).fill('ResourceLimitExceeded')); expect(result.actualNodes).toBe(result.nodes); expect(result.gridCells).toBeGreaterThan(0); expect(result.shadowPixels).toBeGreaterThan(0);
    const identities = [...new Set(result.objects.map(object => object.nodeId))]; expect(identities).toHaveLength(3); expect(identities).toEqual(result.expectedOrder); expect(result.canonicalIds).toBe(0);
    expect(result.fragments.map(fragment => fragment.nodeId)).toEqual(identities.flatMap(id => [id, id, id, id]));
    for (const id of identities) {
      const parts = result.objects.filter(object => object.nodeId === id); expect(parts.map(part => part.pane)).toEqual(['corner', 'frozenRows', 'frozenColumns', 'body']);
      expect(result.fragments.filter(fragment => fragment.nodeId === id).map(fragment => fragment.hidden)).toEqual([null, 'true', 'true', 'true']);
      for (const part of parts) {
        const fragment = result.fragments.find(fragment => fragment.nodeId === id && fragment.pane === part.pane)!;
        // CSSOM serializes point lengths to about six significant digits.
        expect(Math.abs(fragment.x - (rtl ? 420 - part.clip.x - part.clip.width : part.clip.x))).toBeLessThan(.001); expect(Math.abs(fragment.width - part.clip.width)).toBeLessThan(.001);
      }
      for (const part of parts) for (const projection of part.projections ?? []) { expect(projection.owned).toBe(true); expect(projection.to - projection.from).toBe(projection.text.length); }
      expect(parts.map(part => part.text)).toEqual(Array(4).fill(parts[0]!.text));
    }
    if (input.name.endsWith('08-images-shapes')) {
      expect(result.imageCount).toBe(4); expect(result.imageUrls).toBe(1); expect(result.assets.count).toBe(1); expect(result.live).toBe(1); expect(result.unavailable).toBe(0);
      expect(result.cropError).toBe('ResourceLimitExceeded');
    } else { expect(result.assets.count).toBe(0); expect(result.unavailable).toBe(4); expect(result.staleDiagnostics).toBeGreaterThan(0); }
  }
});


test('offscreen object anchors still project visible hanging text outside the anchor', async ({ page }) => {
  const input = inputs.find(input => input.name === '08-images-shapes')!;
  const result = await page.evaluate(files => window.workbookDrawingWindow.drawingWindowMeasurement(files, false, true), input.files);
  expect(result.overflowParts.length).toBeGreaterThan(0);
  expect(result.overflowParts.every(part => part.x > 420 && part.left < -80)).toBe(true);
  expect(result.overflowParts.some(part => part.visible)).toBe(true);
  expect(result.diagnostics).toContain('drawing-text-overflow');
  expect(result).toMatchObject({ unchanged: true, preserved: true, released: true, cleaned: true });
});
