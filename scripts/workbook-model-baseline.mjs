import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { mkdir, writeFile } from 'node:fs/promises';
import { createWorkbook, openDocument } from '@dagent-office/core';
import { baselineRows, baselineColumns, generateBlock } from './workbook-baseline-fixture.mjs';

// Public HTML package model probe. Native XLSX and calculation are separate gates.
const encoder = new TextEncoder(), decoder = new TextDecoder();
const files = new Map(createWorkbook({ documentId: 'model-baseline' }));
const manifest = JSON.parse(decoder.decode(files.get('manifest.json'))), metadata = JSON.parse(decoder.decode(files.get('metadata.json')));
metadata.workbook.activeSheetId = 'baseline'; metadata.workbook.sheets[0].sheetId = 'baseline';
files.set('metadata.json', encoder.encode(JSON.stringify(metadata)));
const digest = createHash('sha256');
for (let row = 0; row < baselineRows; row += 256) {
  const bytes = generateBlock(row), path = 'sheets/baseline/' + row + '.html';
  digest.update(bytes); files.set(path, bytes);
  manifest.chunks.push({ sheetId: 'baseline', row, column: 0, path, sha256: createHash('sha256').update(bytes).digest('hex') });
}
files.set('manifest.json', encoder.encode(JSON.stringify(manifest)));
let start = performance.now(); const model = openDocument(files), openMs = performance.now() - start;
assert.equal(model.kind, 'workbook'); assert.equal('html' in model.snapshot(), false);
const snapshotBytes = encoder.encode(JSON.stringify(model.snapshot())).byteLength;
assert.ok(snapshotBytes < 20_000);
let cells = 0, formulas = 0; start = performance.now();
for (let row = 0; row < baselineRows; row += 128) {
  const window = model.readWindow({ sheetId: 'baseline', row, column: 0, rows: Math.min(128, baselineRows - row), columns: baselineColumns });
  cells += window.length; formulas += window.filter(cell => cell.formula !== undefined).length;
}
const scanMs = performance.now() - start;
assert.equal(cells, 200_000); assert.equal(formulas, 20_000);
assert.deepEqual(model.getCell('baseline', 9999, 19).value, { type: 'number', value: 3_400_152 });
assert.equal(model.getCell('baseline', 9999, 19).formula.cacheState, 'imported');
const report = { kind: 'R3-03-public-model-probe', generatedAt: new Date().toISOString(), node: process.version,
  rows: baselineRows, columns: baselineColumns, cellCount: cells, formulaCount: formulas, sourceSha256: digest.digest('hex'),
  openMs, allWindowsScanMs: scanMs, snapshotBytes, processLifetimePeakRssBytes: process.resourceUsage().maxRSS * 1024,
  limitations: ['HTML package reading only; not native XLSX or calculation acceptance.', 'Local environment; final fixed-resource/browser gates remain outstanding.'] };
await mkdir('.cache/workbook-baseline', { recursive: true });
await writeFile('.cache/workbook-baseline/model.json', JSON.stringify(report, null, 2) + '\n');
console.log(JSON.stringify(report, null, 2));
