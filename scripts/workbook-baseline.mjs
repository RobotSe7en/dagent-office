import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { availableParallelism, totalmem, platform, arch } from 'node:os';
import { Worker, isMainThread, parentPort } from 'node:worker_threads';
import { indexWorkbookChunk } from '../packages/core/dist/workbook/chunk-index.js';
import { createFormulaTemplatePool } from '../packages/core/dist/workbook/formula.js';
import { baselineRows, baselineColumns, generateBlock } from './workbook-baseline-fixture.mjs';

// This measures source storage only. It does not certify native XLSX opening,
// general formula parsing/evaluation, browser rendering or transaction history.
if (isMainThread) {
  const before = process.memoryUsage().rss, delay = [], start = performance.now();
  let previous = start, peak = before;
  const timer = setInterval(() => {
    const now = performance.now(); delay.push(Math.max(0, now - previous - 10)); previous = now;
    peak = Math.max(peak, process.memoryUsage().rss);
  }, 10);
  const worker = new Worker(new URL(import.meta.url), { resourceLimits: { maxOldGenerationSizeMb: 1536 } });
  try {
    const result = await new Promise((resolve, reject) => {
      worker.on('message', message => { if (message.kind === 'result') resolve(message.result); });
      worker.on('error', reject); worker.on('exit', code => reject(new Error('Baseline Worker exited before its result: ' + code)));
    });
    clearInterval(timer); delay.sort((a, b) => a - b);
    const report = { kind: 'R3-02-source-storage-probe', generatedAt: new Date().toISOString(),
      environment: { node: process.version, platform: platform(), arch: arch(), availableCpus: availableParallelism(), systemMemoryBytes: totalmem() },
      ...result, elapsedMs: performance.now() - start, parentEventLoopDelayP95Ms: delay[Math.floor(delay.length * .95)] ?? 0,
      processIncrementalPeakRssBytes: Math.max(0, peak - before),
      processLifetimePeakRssBytes: process.resourceUsage().maxRSS * 1024,
      limitations: ['Source storage probe; not XLSX open or calculation acceptance.', 'Memory includes Node and its Worker; browser measurements are separate.', 'Formula subset parsed and interned; cache not evaluated.'] };
    assert.equal(report.cellCount, baselineRows * baselineColumns); assert.equal(report.formulaCount, baselineRows * 2); assert.equal(report.formulaTemplates, 2);
    assert.ok(report.processIncrementalPeakRssBytes < 2 * 1024 ** 3);
    await mkdir('.cache/workbook-baseline', { recursive: true });
    await writeFile('.cache/workbook-baseline/node.json', JSON.stringify(report, null, 2) + '\n');
    console.log(JSON.stringify(report, null, 2));
  } finally { clearInterval(timer); await worker.terminate(); }
} else {
  const chunks = [], hash = createHash('sha256'), formulas = createFormulaTemplatePool();
  let generationMs = 0, indexingMs = 0;
  for (let row = 0; row < baselineRows; row += 256) {
    let start = performance.now(); const bytes = generateBlock(row); generationMs += performance.now() - start;
    hash.update(bytes); start = performance.now();
    chunks.push(indexWorkbookChunk(bytes, { sheetId: 'baseline', row, column: 0 }, undefined, { formulas })); indexingMs += performance.now() - start;
  }
  const windows = [], start = performance.now();
  for (const row of [0, 255, 256, baselineRows - 1]) {
    const cells = chunks[Math.floor(row / 256)].readWindow({ row, column: 0, rows: 1, columns: 20 });
    assert.equal(cells.length, 20); assert.ok(cells[1].html.includes('客户 ' + row + ' 😀禰󠄀'));
    assert.ok(cells[18].html.includes('=SUM(C' + (row + 1) + ':R' + (row + 1) + ')'));
    windows.push(cells);
  }
  const queryMs = performance.now() - start;
  assert.equal(chunks[0].getFormula(0, 18), chunks.at(-1).getFormula(baselineRows - 1, 18));
  // Worker replies carry requested cells only; the source/index stays here.
  parentPort.postMessage({ kind: 'window', cells: windows[0] });
  parentPort.postMessage({ kind: 'result', result: {
    rows: baselineRows, columns: baselineColumns, chunks: chunks.length, sourceSha256: hash.digest('hex'),
    cellCount: chunks.reduce((sum, chunk) => sum + chunk.cellCount, 0), formulaCount: chunks.reduce((sum, chunk) => sum + chunk.formulaCount, 0),
    sourceBytes: chunks.reduce((sum, chunk) => sum + chunk.sourceBytes, 0), indexBytes: chunks.reduce((sum, chunk) => sum + chunk.indexBytes, 0),
    formulaTemplates: formulas.size, formulaTemplateKeyCodeUnits: formulas.keyCodeUnits,
    generationMs, indexingMs, fourWindowsQueryMs: queryMs, transferredWindowCells: windows[0].length,
  } });
}
