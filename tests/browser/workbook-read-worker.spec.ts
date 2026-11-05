import { expect, test } from '@playwright/test';
import { readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import type * as fixture from './fixtures/workbook-read-host.js';
declare global { interface Window { workbookReadHost: typeof fixture } }
let workerCode = '', hostCode = '';
test.beforeAll(async () => {
  async function bundle(name: string): Promise<string> {
    const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL(`./fixtures/${name}.ts`, import.meta.url)), formats: ['iife'], name: 'WorkbookReadProbe' } } });
    if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected read-worker bundle');
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
    return output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  }
  workerCode = await bundle('workbook-read-worker'); hostCode = await bundle('workbook-read-host');
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: hostCode }); });
test('native XLSX pictures and charts stay in a host Worker and feed the public preview through bounded reads', async ({ page }) => {
  for (const name of ['07-three-charts', '08-images-shapes']) {
    const result = await page.evaluate(({ code, bytes, name }) => window.workbookReadHost.nativeWorkerRead(code, bytes, name, false), { code: workerCode, bytes: [...readFileSync(`examples/xlsx/${name}.xlsx`)], name });
    expect(result).toMatchObject({ frozen: true, retained: true, detachedInput: true, failures: [] });
    expect(['InvalidOperation', 'OperationCancelled']).toContain(result.closed); expect(result.progress).toBeGreaterThan(0);
    expect(result.first.objectFragmentCount).toBeGreaterThan(0); expect(result.first.documentId).toBe(name); expect(result.largestReplyUnits).toBeLessThan(1024 * 1024);
  }
});
test('10,000 actual XLSX rows open and scroll through the production read transport without moving full cell chunks to the main thread', async ({ page }, info) => {
  test.setTimeout(120_000);
  const bytes = [...readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx')];
  const result = await page.evaluate(({ code, bytes }) => window.workbookReadHost.nativeWorkerRead(code, bytes, 'ten-thousand-native', true), { code: workerCode, bytes });
  expect(result.counts).toEqual({ kind: 'audit', cells: 200_000, formulas: 20_000, chunks: 40 });
  expect(result.last).toHaveLength(20); expect(result.last.at(-1)?.value).toEqual({ type: 'number', value: 3_400_152 });
  expect(result).toMatchObject({ frozen: true, retained: true, detachedInput: true, failures: [] });
  expect(result.last.filter(cell => cell.formula?.cacheState === 'imported')).toHaveLength(2);
  expect(result.afterScroll.formulaCaches.imported).toBeGreaterThan(0); expect(result.afterScroll.viewport.scrollTopPt).toBe(149_800);
  expect(result.first.storedCellCount).toBeLessThan(5000); expect(result.afterScroll.storedCellCount).toBeLessThan(5000);
  expect(result.snapshotUnits).toBeLessThan(100_000); expect(result.largestReplyUnits).toBeLessThan(1024 * 1024); expect(result.ticks).toBeGreaterThan(10);
  const cancelled = await page.evaluate(({ code, bytes }) => window.workbookReadHost.cancelNativeWorker(code, bytes), { code: workerCode, bytes });
  expect(cancelled.error).toBe('OperationCancelled'); expect(cancelled.progress).toBeGreaterThan(0);
  // Measurements are recorded; this host is not the final pinned 4-vCPU/8-GiB
  // acceptance environment, so a passing test does not certify that SLA.
  const report = JSON.stringify({ ...result, cancelled }, null, 2);
  writeFileSync(info.outputPath('native-workbook-worker.json'), report + '\n');
  await info.attach('native-workbook-worker.json', { body: report, contentType: 'application/json' });
});
