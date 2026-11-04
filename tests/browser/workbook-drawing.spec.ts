import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import type { WorkbookSnapshot } from '@dagent-office/core';
import { readXlsxRecords } from '../../packages/office/src/xlsx/read.js';
import { projectXlsxRecords } from '../../packages/office/tests/helpers/xlsx-record-model.js';
import type { measureContentSource, rejectContentSource } from './fixtures/workbook-content-source.js';

declare global { interface Window { workbookContentHost: { measureContentSource: typeof measureContentSource; rejectContentSource: typeof rejectContentSource } } }
const inputs: { name: string; snapshot: WorkbookSnapshot; text: Record<string, string> }[] = [];
let code = '';
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-content-source.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing fixture bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes', '09-classic-notes', '10-modern-threads']) {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`))), snapshot = model.snapshot();
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    inputs.push({ name: prefix + name, snapshot, text: Object.fromEntries(ids.map(id => [id, model.getNode(id)!.text])) });
  }
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: code }); });

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
