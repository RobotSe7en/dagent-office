import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { importXlsx } from '../../packages/office/src/xlsx-import.js';
import type { drawingWindowMeasurement } from './fixtures/workbook-drawing-window.js';
import type { annotationWindowMeasurement, commentTextMeasurement, windowControllerMeasurement } from './fixtures/workbook-window-controller.js';
import type { commentPanelMeasurement } from './fixtures/workbook-comment-panel.js';

declare global { interface Window { workbookDrawingWindow: { drawingWindowMeasurement: typeof drawingWindowMeasurement; windowControllerMeasurement: typeof windowControllerMeasurement; annotationWindowMeasurement: typeof annotationWindowMeasurement; commentTextMeasurement: typeof commentTextMeasurement; commentPanelMeasurement: typeof commentPanelMeasurement } } }
let code = '';
const inputs: { name: string; files: [string, number[]][] }[] = [];
const comments: typeof inputs = [];
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-drawing-window.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingWindow' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing-window bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`), { documentId: 'drawing-window' });
    inputs.push({ name: prefix + name, files: [...files].map(([path, bytes]) => [path, [...bytes]]) });
  }
  for (const prefix of ['', 'producer/']) for (const name of ['09-classic-notes', '10-modern-threads']) {
    const files = await importXlsx(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`), { documentId: 'comment-window' });
    comments.push({ name: prefix + name, files: [...files].map(([path, bytes]) => [path, [...bytes]]) });
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

test('whole workbook windows commit atomically across scrolling, sheet changes, late providers and disposal', async ({ page }) => {
  for (const input of inputs.filter(input => !input.name.includes('/'))) {
    const result = await page.evaluate(files => window.workbookDrawingWindow.windowControllerMeasurement(files), input.files);
    expect(result, input.name).toMatchObject({ first: true, errors: ['RevisionConflict', 'InvalidDocument'], retainedWhilePending: true, switched: true,
      obsoleteCancelled: true, empty: { sheetId: 'controller-empty', objects: 0, width: 380, height: 260 }, lateCellIgnored: true, releasedOld: true,
      budget: 'ResourceLimitExceeded', failedMountEmpty: true, cancelledOnDispose: true, released: true, imagesCleared: true, emptyAfterDispose: true, formatterStillOwned: true, unhandled: [], unchanged: true });
    expect(result.nodes).toBe(result.actualNodes);
    expect(result.races).toHaveLength(2);
    for (const race of result.races) expect(race).toMatchObject({ cancelled: true, aborted: true, preserved: true });
  }
});

test('native classic notes and modern threads retain source content and pane locations in complete windows', async ({ page }) => {
  for (const input of comments) {
    const result = await page.evaluate(files => window.workbookDrawingWindow.annotationWindowMeasurement(files), input.files);
    expect(result.visible.map(item => item.id), input.name).toEqual(result.expected);
    // The pinned Calc producer drops modern threads; its output contains none.
    expect(result.visible.length).toBe(input.name === 'producer/10-modern-threads' ? 0 : input.name === '10-modern-threads' ? 2 : 1);
    expect(result.recovery).toEqual(result.expectedRecovery); expect(result.unchanged).toBe(true);
    for (const item of result.visible) { expect(item.body?.length).toBeGreaterThan(0); expect(item.ranges.length).toBeGreaterThan(0); }
    // These independent sources have no frozen comments: scrolling to the
    // last row/column must remove their geometry from the published window.
    expect(result.offscreen).toBe(0);
  }
});

test('comment body flow preserves native run styles, paragraph offsets and UTF-16 selection without source identities', async ({ page }) => {
  for (const input of comments.filter(input => input.name !== 'producer/10-modern-threads')) for (const rich of input.name === '09-classic-notes' ? [false, true] : [false]) {
    const result = await page.evaluate(({ files, rich }) => window.workbookDrawingWindow.commentTextMeasurement(files, rich), { files: input.files, rich });
    expect(result.unchanged).toBe(true); expect(result.results.length).toBeGreaterThan(0);
    for (const body of result.results) {
      expect(body).toMatchObject({ detached: true, unchangedSource: true, canonicalIds: 0, selected: true, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded'] });
      expect(body.nodes).toBe(body.actualNodes); expect(body.fonts).toBeGreaterThan(0);
      for (const piece of body.offsets) { expect(piece.text).toBe(piece.expected); expect(piece.to - piece.from).toBe(piece.text.length); expect(piece.sourceNodeId).toBe(body.id); }
      if (!rich) expect(body.text).toBe(body.expected);
      else {
        const start = body.expected.indexOf('混排😀禰󠄀'), runs = body.formats.filter(run => run.from >= start && run.from < start + '混排😀禰󠄀'.length);
        expect(runs.map(run => run.text).join('')).toBe('混排😀禰󠄀');
        for (const run of runs) expect(run).toMatchObject({ underline: 'underline', style: 'double', strike: 'line-through', color: 'rgb(255, 0, 0)' });
        expect(body.formats.find(run => run.text === 'No underline')).toMatchObject({ underline: 'none', strike: '' });
      }
    }
  }
});

test('read-only comment candidates retain recovery state and reply parents while failed replacements preserve the active panel', async ({ page }) => {
  for (const name of ['09-classic-notes', '10-modern-threads']) for (const recovered of name === '10-modern-threads' ? [false, true] : [false]) {
    const input = comments.find(input => input.name === name)!;
    const result = await page.evaluate(({ files, recovered }) => window.workbookDrawingWindow.commentPanelMeasurement(files, recovered), { files: input.files, recovered });
    expect(result).toMatchObject({ detached: true, editable: 0, released: true, empty: true, unchanged: true, plain: { elements: 0, bodies: 0 } });
    expect(result.nodes).toBe(result.actualNodes); expect(result.fonts).toBeGreaterThan(0);
    expect(result.errors).toEqual([...Array(recovered ? 4 : 3).fill('ResourceLimitExceeded'), 'RevisionConflict', 'InvalidOperation']);
    expect(result.plain.text).toContain('<script>author</script>'); expect(result.plain.text).toContain('<img src="https://invalid.example/private">仅作为文字');
    for (const body of result.bodies) { expect(body.text?.length).toBeGreaterThan(0); expect(body.projections).toBeGreaterThan(0); }
    if (recovered) { expect(result.warning).toContain('位置记录不一致'); expect(result.warning).toContain('原状态：已处理'); expect(result.replyParents).toHaveLength(1); expect(result.text).toContain('回复中的纯文本'); }
    else expect(result.warning).toBeUndefined();
  }
});
