import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type * as fixture from './fixtures/workbook-preview.js';
declare global { interface Window { workbookPreview: typeof fixture } }
let code = '';
const inputs: { name: string; files: [string, number[]][] }[] = [];
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-preview.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookPreview' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected preview bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const name of ['09-classic-notes', '10-modern-threads', '08-images-shapes']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${name}.xlsx`), { documentId: 'public-preview' });
    inputs.push({ name, files: [...files].map(([path, bytes]) => [path, [...bytes]]) });
  }
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: code }); });
for (const rtl of [false, true]) test(`public workbook preview selects cells and switches sheets without serializing source in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  for (const input of inputs.filter(input => input.name !== '08-images-shapes')) {
    const start = await page.evaluate(({ files, rtl }) => window.workbookPreview.mountPublicPreview(files, rtl), { files: input.files, rtl });
    expect(start).toMatchObject({ unchanged: true, serializations: 0, frames: 1, errors: [], report: { kind: 'workbook', layout: 'grid' } });
    await page.mouse.click(start.x, start.y);
    expect(await page.evaluate(() => window.workbookPreview.previewInspect())).toMatchObject({ address: 'A1', selection: { documentId: 'public-preview', baseRevision: 0, anchor: { kind: 'cells', ranges: [{ from: { row: 0, column: 0 }, to: { row: 0, column: 0 } }] } } });
    await page.evaluate(() => window.workbookPreview.previewAction('comment'));
    const open = await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { fonts: unknown[] } };
    expect(open.report.fonts.length).toBeGreaterThan(0);
    await page.evaluate(() => window.workbookPreview.previewAction('close'));
    const empty = await page.evaluate(() => window.workbookPreview.previewAction('empty'));
    expect(empty).toMatchObject({ report: { sheetId: 'preview-empty', storedCellCount: 0 }, address: '', frames: 1 });
    expect(await page.evaluate(() => window.workbookPreview.previewAction('bad-sheet'))).toMatchObject({ report: { sheetId: 'preview-empty' }, errors: ['InvalidOperation'], frames: 1 });
    await page.evaluate(() => window.workbookPreview.previewAction('first'));
    await page.mouse.move(start.x + (rtl ? -50 : 50), start.y + 80); await page.mouse.wheel(0, 200);
    await expect.poll(async () => (await page.evaluate(() => window.workbookPreview.previewInspect()) as { report: { viewport: { scrollTopPt: number } } }).report.viewport.scrollTopPt).toBeGreaterThan(0);
    expect(await page.evaluate(() => window.workbookPreview.previewAction('dispose'))).toEqual({ removed: true, frames: 0, unchanged: true, serializations: 0, fieldsCleared: true });
  }
});
test('public workbook preview retains accepted data on failure and settles superseded and disposed providers', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.publicPreviewLifecycle(files), inputs.find(input => input.name === '08-images-shapes')!.files);
  expect(result).toEqual({ scrollRace: true, failed: 'RevisionConflict', retained: true, updateRace: true, disposed: 'InvalidOperation', errors: ['RevisionConflict'], unhandled: [], providerAlive: true, released: true, removed: true });
});
test('control construction is atomic and open comments and selections share the complete window budget', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.previewDecorationMeasurement(files), inputs.find(input => input.name === '09-classic-notes')!.files);
  expect(result).toMatchObject({ errors: ['UnsupportedFeature', 'ResourceLimitExceeded', 'cancelled'], highlights: 1, releasedCandidates: 2, releasedAll: 3, empty: true });
  expect(result.nodes).toBe(result.actualNodes); expect(result.fonts).toBeGreaterThan(0); expect(result.panel?.length).toBeGreaterThan(0);
});
test('initial failures release their candidate and a host retry can satisfy the same readiness promise', async ({ page }) => {
  const result = await page.evaluate(files => window.workbookPreview.initialPreviewFailure(files), inputs.find(input => input.name === '09-classic-notes')!.files);
  expect(result).toEqual({ rejected: 'InvalidOperation', frames: 0, errors: ['InvalidOperation', 'InvalidOperation'], retried: true, finalFrames: 1 });
});
