import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import type { WorkbookSnapshot } from '@dagent-office/core';
import { readXlsxRecords } from '../../packages/office/src/xlsx/read.js';
import { projectXlsxRecords } from '../../packages/office/tests/helpers/xlsx-record-model.js';
import type { measureContentSource, rejectContentSource } from './fixtures/workbook-content-source.js';
import type { measureImageView, disposeImageView } from './fixtures/workbook-image-view.js';
import type { measureObjectGeometry } from './fixtures/workbook-object-geometry.js';

declare global { interface Window { workbookContentHost: { measureContentSource: typeof measureContentSource; rejectContentSource: typeof rejectContentSource; measureImageView: typeof measureImageView; disposeImageView: typeof disposeImageView; measureObjectGeometry: typeof measureObjectGeometry } } }
const inputs: { name: string; snapshot: WorkbookSnapshot; text: Record<string, string>; resources: Record<string, string> }[] = [];
let code = '';
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-content-source.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing fixture bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes', '09-classic-notes', '10-modern-threads']) {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`))), snapshot = model.snapshot();
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    inputs.push({ name: prefix + name, snapshot, text: Object.fromEntries(ids.map(id => [id, model.getNode(id)!.text])), resources: Object.fromEntries(snapshot.manifest.assets.map(asset => [asset.path, Buffer.from(model.getAsset(asset.path)!).toString('base64')])) });
  }
});

for (const rtl of [false, true]) test(`cropped worksheet pictures retain full rotated and flipped geometry across four ${rtl ? 'RTL' : 'LTR'} panes`, async ({ page }, info) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(({ snapshot, resources, rtl }) => window.workbookContentHost.measureImageView(snapshot, resources, rtl), { ...input, rtl });
    expect(result).toMatchObject({ invalid: 'InvalidDocument', cancelled: 'OperationCancelled', connected: 'InvalidOperation', kept: true, cleared: true, badGeometry: ['InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded'], extremeAccepted: true, peak: 8, decoded: 17 });
    const [cropX, cropY] = input.name.startsWith('producer/') ? [.10035, .05039] : [.1, .05];
    expect(result.panes).toHaveLength(4); expect(result.sourceRotation).toBe(15); expect(result.sourceCrops).toEqual([cropX, cropY, cropX, cropY]);
    const cosine = Math.cos(Math.PI / 12), sine = Math.sin(Math.PI / 12);
    for (const pane of result.panes) {
      expect(pane.body.slice(2)).toEqual([200, 150]); expect(pane.natural).toEqual([32, 24]);
      for (const [index, value] of [-cosine, -sine, -sine, cosine].entries()) expect(pane.matrix[index]).toBeCloseTo(value, 5);
      expect(pane.transformed[0]).toBeCloseTo(200 * cosine + 150 * sine, 1); expect(pane.transformed[1]).toBeCloseTo(200 * sine + 150 * cosine, 1);
      for (const [index, value] of [-cropX! / (1 - 2 * cropX!) * 100, -cropY! / (1 - 2 * cropY!) * 100, 100 / (1 - 2 * cropX!), 100 / (1 - 2 * cropY!)].entries()) expect(pane.crop[index]).toBeCloseTo(value, 3);
      const sourceLeft = pane.pane === 'body' || pane.pane === 'frozenRows' ? -10 : 10;
      expect(pane.clip[0]! + pane.body[0]!).toBeCloseTo(rtl ? 240 - sourceLeft - 200 : sourceLeft, 0);
      expect(pane.clip[1]! + pane.body[1]!).toBeCloseTo(pane.pane === 'body' || pane.pane === 'frozenColumns' ? -5 : 5, 0);
    }
    await page.locator('#workbook-image-probe').screenshot({ path: info.outputPath(`${input.name.replace('/', '-')}-${rtl ? 'rtl' : 'ltr'}.png`) });
    await page.evaluate(() => window.workbookContentHost.disposeImageView());
    await expect(page.locator('#workbook-image-probe')).toHaveCount(0);
  }
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: code }); });

test('workbook geometry shares preset paths, native colors and bounded gradient/arrow primitives', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureObjectGeometry(snapshot), input.snapshot);
    expect(result.results.map(object => object.preset)).toEqual(['rect', 'roundRect', 'rect']);
    for (const object of result.results.slice(1)) {
      expect(object.paths).toHaveLength(1);
      expect(object.paths[0]).toMatchObject({ fill: 'rgba(232,240,230,1)', stroke: 'rgba(40,116,84,1)', width: '1' });
      expect(object.textRect!.width).toBeGreaterThan(0); expect(object.textRect!.height).toBeGreaterThan(0);
    }
    expect(result.errors).toEqual(Array(3).fill('ResourceLimitExceeded')); expect(result.nodes).toBe(19);
    expect(new Set(result.ids).size).toBe(result.ids.length); expect(new Set(result.arrows).size).toBe(4);
    expect(result.arrows.every(reference => result.ids.includes(reference!.slice(5, -1)))).toBe(true);
    expect(result.stops).toEqual(['rgba(255,0,0,1)', 'rgba(0,0,255,1)']);
  }
});

test('bounded auxiliary projections preserve native drawing and comment text without activating image paths', async ({ page }) => {
  const requests: string[] = []; page.on('request', request => requests.push(request.url()));
  for (const input of inputs) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureContentSource(snapshot), input.snapshot);
    expect(result.rootTexts, input.name).toEqual(input.text);
    expect(result, input.name).toMatchObject({ isolated: true, atomic: true, secondSurvives: true, released: true, limited: input.snapshot.content ? 'ResourceLimitExceeded' : '', unknown: 'InvalidOperation', duplicated: 'InvalidOperation', disposed: 'InvalidOperation' });
    expect(result.imageAttributes.every(image => image.src === null && input.snapshot.manifest.assets.some(asset => asset.path === image.path))).toBe(true);
  }
  await page.waitForTimeout(50); expect(requests).toEqual([]);
  await expect(page.locator('[data-office-view="workbook-content-measurement"]')).toHaveCount(0);
});

test('auxiliary parsing rejects active content, foreign nodes, repeated identities and budgets while inert', async ({ page }) => {
  const snapshot = inputs.find(input => input.name === '08-images-shapes')!.snapshot;
  const requests: string[] = []; page.on('request', request => requests.push(request.url()));
  for (const edit of [
    (html: string) => html.replace('<img ', '<img onload="window.injected=true" '),
    (html: string) => html.replace('<img ', '<svg><image '),
    (html: string) => html.replace(/src="[^"]+"/, 'src="https://example.test/private.png"'),
    (html: string) => html.replace('</section>', '<script data-node-id="injection">window.injected=true</script></section>'),
    (html: string) => html.replace('</section>', '<p data-node-id="injection">unowned</p></section>'),
    (html: string) => html.replace(/data-node-id="[^"]+"/g, 'data-node-id="duplicate"'),
  ]) {
    const broken = { ...snapshot, content: { ...snapshot.content!, html: edit(snapshot.content!.html) } };
    expect(await page.evaluate(snapshot => window.workbookContentHost.rejectContentSource(snapshot), broken)).toBe('InvalidDocument');
  }
  for (const limits of [[1, 8 * 1024 * 1024], [100_000, 1]]) expect(await page.evaluate(({ snapshot, limits }) => window.workbookContentHost.rejectContentSource(snapshot, limits[0], limits[1]), { snapshot, limits })).toBe('ResourceLimitExceeded');
  expect(await page.evaluate(snapshot => window.workbookContentHost.rejectContentSource(snapshot, 0), snapshot)).toBe('InvalidOperation');
  await page.waitForTimeout(50); expect(requests).toEqual([]);
});
