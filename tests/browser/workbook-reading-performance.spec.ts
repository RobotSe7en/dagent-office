import { expect, test } from '@playwright/test';
import { build } from 'vite';
import { createHash } from 'node:crypto';
import { appendFileSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import type * as fixture from './fixtures/workbook-reading-budget.js';
declare global { interface Window { workbookBudget: typeof fixture } }
const p95 = (samples: number[]) => [...samples].sort((a, b) => a - b)[Math.ceil(samples.length * .95) - 1]!;
// DOM snapshots and action screenshots in functional traces materially delay
// native input. Keep performance samples free of that test instrumentation;
// failures still retain the full JSON samples and process/resource report.
test.use({ trace: 'off', screenshot: 'off', video: 'off' });
test('pinned-resource native reading, actual selection and wheel feedback, DOM and cancellation', async ({ page, browserName }, info) => {
  test.skip(process.env.DAGENT_WORKBOOK_BUDGET !== '1', 'Run through the explicit 4-vCPU/8-GiB acceptance wrapper.');
  test.setTimeout(240_000);
  const bundle = async (name: string) => {
    const result = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/' + name + '.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookBudget' } } });
    if (!Array.isArray(result) && !('output' in result)) throw Error('Unexpected performance fixture bundle.');
    const output = Array.isArray(result) ? result.flatMap(part => part.output) : result.output;
    return output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  };
  const worker = await bundle('workbook-read-worker'), host = await bundle('workbook-reading-budget');
  await page.route('**/workbook-reading-budget.html', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><html><head><meta charset="utf-8"><title>Workbook reading budget</title></head><body></body></html>' }));
  await page.goto('/workbook-reading-budget.html');
  await page.evaluate(code => { const script = document.createElement('script'); script.textContent = code; document.head.append(script); }, host);
  const bytes = readFileSync('examples/xlsx/15-ten-thousand-rows.xlsx');
  try {
    await page.evaluate(({ worker, bytes }) => window.workbookBudget.openBudget(worker, bytes), { worker, bytes: [...bytes] });
    const profile = process.env.DAGENT_WORKBOOK_CPU_PROFILE && browserName === 'chromium' ? await page.context().newCDPSession(page) : undefined;
    if (profile) {
      await profile.send('Profiler.enable'); await profile.send('Profiler.start');
      await profile.send('Tracing.start', { categories: 'devtools.timeline,disabled-by-default-devtools.timeline.frame,blink.user_timing,input,latencyInfo', transferMode: 'ReturnAsStream' });
    }
    for (let index = 0; index < 30; index++) {
      const point = await page.evaluate(index => window.workbookBudget.budgetPoint(index), index); await page.mouse.click(point.x, point.y);
      await page.evaluate(index => window.workbookBudget.budgetSample('selection', index), index);
    }
    for (let index = 0; index < 40; index++) {
      const point = await page.evaluate(() => window.workbookBudget.budgetPoint(3)); await page.mouse.move(point.x, point.y); await page.mouse.wheel(0, index < 20 ? 80 : -80);
      await page.evaluate(index => window.workbookBudget.budgetSample('wheel', index), index);
    }
    if (profile) {
      writeFileSync(info.outputPath('workbook-wheel.cpuprofile'), JSON.stringify((await profile.send('Profiler.stop')).profile));
      const complete = new Promise<{ stream?: string }>(resolve => profile.once('Tracing.tracingComplete', resolve)); await profile.send('Tracing.end');
      const { stream } = await complete; if (!stream) throw Error('The browser did not return its diagnostic trace.');
      const path = info.outputPath('workbook-wheel-trace.json'); writeFileSync(path, ''); let size = 0;
      try { for (;;) { const part = await profile.send('IO.read', { handle: stream, size: 1024 * 1024 }); const bytes = Buffer.from(part.data, part.base64Encoded ? 'base64' : 'utf8');
        size += bytes.length; if (size > 32 * 1024 * 1024) throw Error('The diagnostic trace exceeded its output budget.'); appendFileSync(path, bytes); if (part.eof) break;
      } } finally { await profile.send('IO.close', { handle: stream }); await profile.detach(); }
    }
    const measured = await page.evaluate(() => window.workbookBudget.budgetState()) as { firstScreenMs: number; maxFragments: number; selections: number[]; wheels: number[]; errors: unknown[]; [key: string]: unknown };
    const counts = await page.evaluate(() => window.workbookBudget.budgetAudit());
    const cancellation = await page.evaluate(() => window.workbookBudget.budgetCancel());
    const sha256 = (value: Uint8Array | string) => createHash('sha256').update(value).digest('hex');
    const report = { ...measured, sourceSha256: sha256(bytes), sourceBytes: bytes.length, workerBundleSha256: sha256(worker), hostBundleSha256: sha256(host), viewport: { width: 1280, height: 800 }, grid: { widthPt: 880, heightPt: 450 }, counts,
      selectionP95Ms: p95(measured.selections), wheelP95Ms: p95(measured.wheels), cancellation,
      instrumentation: { host: 'standalone public-SDK preview', trace: 'off', screenshots: 'off', video: 'off', cpuProfiler: Boolean(profile) } };
    writeFileSync(info.outputPath('workbook-reading-budget.json'), JSON.stringify(report, null, 2) + '\n');
    await info.attach('workbook-reading-budget.json', { body: JSON.stringify(report), contentType: 'application/json' });
    expect(report.counts).toEqual({ kind: 'audit', cells: 200_000, formulas: 20_000, chunks: 40 });
    expect(report.errors).toEqual([]); expect(Number(measured.initialFragments)).toBeGreaterThan(0); expect(report.maxFragments).toBeLessThanOrEqual(5000); expect(report.firstScreenMs).toBeLessThanOrEqual(30_000);
    expect(report.cancellation).toMatchObject({ error: 'OperationCancelled', retained: true });
    expect(report.selectionP95Ms).toBeLessThanOrEqual(100); expect(report.wheelP95Ms).toBeLessThanOrEqual(100);
  } finally { await page.evaluate(() => window.workbookBudget.closeBudget()); }
});
