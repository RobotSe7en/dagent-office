import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import type { WorkbookSnapshot } from '@dagent-office/core';
import { readXlsxRecords } from '../../packages/office/src/xlsx/read.js';
import { readDocumentZip, writeOpcZip } from '../../packages/office/src/document-zip.js';
import { projectXlsxRecords } from '../../packages/office/tests/helpers/xlsx-record-model.js';
import type { measureContentSource, rejectContentSource } from './fixtures/workbook-content-source.js';
import type { measureImageView, disposeImageView } from './fixtures/workbook-image-view.js';
import type { measureObjectGeometry } from './fixtures/workbook-object-geometry.js';
import type { measureDrawingGradients } from './fixtures/drawing-gradient.js';
import type { measureDrawingText, measureDrawingTextLayout, measureNativeDrawingText, measureDrawingTextFallback } from './fixtures/workbook-drawing-text.js';

declare global { interface Window { workbookContentHost: { measureContentSource: typeof measureContentSource; rejectContentSource: typeof rejectContentSource; measureImageView: typeof measureImageView; disposeImageView: typeof disposeImageView; measureObjectGeometry: typeof measureObjectGeometry; measureDrawingGradients: typeof measureDrawingGradients; measureDrawingText: typeof measureDrawingText; measureDrawingTextLayout: typeof measureDrawingTextLayout; measureNativeDrawingText: typeof measureNativeDrawingText; measureDrawingTextFallback: typeof measureDrawingTextFallback } } }
const inputs: { name: string; snapshot: WorkbookSnapshot; text: Record<string, string>; resources: Record<string, string> }[] = [];
let code = '';
let nativeText: WorkbookSnapshot;
const textEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-drawing-text-semantics.json', 'utf8')) as { sourceSha256: string; cases: { fontPt: number | null; endPt: number | null; paragraphs: string[]; listStyle: string; observed: { lineAdvancePt: number; words: { text: string; boundsPt: number[] }[] } }[] };
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-content-source.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing fixture bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  const nativeBytes = execFileSync('python3', ['-c', "import importlib.util,sys; s=importlib.util.spec_from_file_location('drawing_text_inputs','scripts/check-xlsx-drawing-text.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))"], { maxBuffer: 16 * 1024 * 1024 });
  expect(createHash('sha256').update(nativeBytes).digest('hex')).toBe(textEvidence.sourceSha256);
  nativeText = projectXlsxRecords(await readXlsxRecords(nativeBytes)).model.snapshot();
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes', '09-classic-notes', '10-modern-threads']) {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`))), snapshot = model.snapshot();
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    inputs.push({ name: prefix + name, snapshot, text: Object.fromEntries(ids.map(id => [id, model.getNode(id)!.text])), resources: Object.fromEntries(snapshot.manifest.assets.map(asset => [asset.path, Buffer.from(model.getAsset(asset.path)!).toString('base64')])) });
  }
  for (const kind of ['default', 'inherited', 'zero']) {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/08-images-shapes.xlsx'))), decoder = new TextDecoder(), encoder = new TextEncoder();
    const path = 'xl/drawings/drawing1.xml', theme = 'xl/theme/theme1.xml';
    let xml = decoder.decode(files.get(path)!).replace('<a:ln w="12700"', kind === 'zero' ? '<a:ln w="0"' : '<a:ln');
    if (kind !== 'default') xml = xml.replace('<xdr:txBody>', '<xdr:style><a:lnRef idx="1"/><a:fillRef idx="0"/><a:effectRef idx="0"/><a:fontRef idx="none"/></xdr:style><xdr:txBody>');
    files.set(path, encoder.encode(xml)); files.set(theme, encoder.encode(decoder.decode(files.get(theme)!).replace('<a:ln w="12700"', '<a:ln w="50800"')));
    const { model } = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files)));
    inputs.push({ name: 'line-' + kind, snapshot: model.snapshot(), text: {}, resources: {} });
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

test('physical DrawingML gradients preserve diagonal stop coverage and the scaled flag in browser pixels', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookContentHost.measureDrawingGradients());
  expect(result.errors).toEqual(Array(4).fill('InvalidDocument'));
  for (const entry of result.cases) {
    expect(entry.units).toBe('userSpaceOnUse');
    const expected = entry.scaled ? [0, 0, 200, 100] : [25, -25, 175, 125];
    entry.coordinates.forEach((value, index) => expect(value).toBeCloseTo(expected[index]!, 8));
    // Expected physical projection at pixel centers, independent of the SVG implementation.
    const coordinates = [[.5, .5], [199.5, .5], [.5, 99.5], [199.5, 99.5], [99.5, 49.5]];
    entry.pixels.forEach((pixel, index) => {
      const [x, y] = coordinates[index]!, t = entry.scaled ? (2 * x! + y!) / 500 : (x! + y!) / 300;
      const expectedPixel = [255 * (1 - t), 0, 255 * t, 255];
      pixel.forEach((channel, index) => expect(Math.abs(channel - expectedPixel[index]!)).toBeLessThanOrEqual(2));
    });
  }
});

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

test('omitted native line widths inherit the theme before the final default, and explicit zero wins', async ({ page }) => {
  for (const [kind, width] of [['default', '.75'], ['inherited', '4'], ['zero', '0']]) {
    const input = inputs.find(input => input.name === 'line-' + kind)!;
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureObjectGeometry(snapshot), input.snapshot);
    expect(Number(result.results[1]!.paths[0]!.width)).toBe(Number(width));
  }
});

test('drawing text keeps native styles and exact UTF-16 projections across paragraphs, spans and styled breaks', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingText(snapshot), input.snapshot);
    for (const object of result.native) {
      expect(object.text).toBe(input.text[object.nodeId]); expect(object.pieces.map(piece => piece.text).join('')).toBe(object.text);
      expect(object).toMatchObject({ unchanged: true, detachedSource: true, inheritedIds: 0 }); expect(object.nodes).toBe(object.actualNodes);
      expect(object.fonts.length).toBeGreaterThan(0);
      // Firefox quantizes the used font size to its subpixel layout grid.
      for (const piece of object.pieces) expect(Math.abs(parseFloat(piece.size) - 14 / .75)).toBeLessThan(.02);
      expect(object.pieces.filter(piece => piece.role === 'east-asia').every(piece => piece.family.includes('Noto Sans CJK SC'))).toBe(true);
    }
    expect(result.pieces.map(piece => piece.text).join('')).toBe('A中󠄀\né😀Tail');
    expect(result.pieces.find(piece => piece.text === '󠄀')).toMatchObject({ role: 'east-asia', from: 2, to: 4, sourceNodeId: 'drawing-r2' });
    expect(result.pieces.find(piece => piece.text === '\n')).toMatchObject({ from: 4, to: 5, sourceNodeId: 'drawing-br', size: '24px' });
    expect(result.pieces.find(piece => piece.text.startsWith('e'))!.weight).toBe('700');
    expect(result.pieces.find(piece => piece.text === 'Tail')).toMatchObject({ from: 11, to: 15, sourceNodeId: 'drawing-p2' });
    expect(result.unchanged).toBe(true); expect(result.exact).toBe(result.nodes);
    expect(result.errors).toEqual(['InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidDocument', 'InvalidOperation']);
  }
});

test('drawing text layout uses native insets, bounded auto-fit and readable glyphs inside reflected objects', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('08-images-shapes'))) {
    const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingTextLayout(snapshot), input.snapshot);
    const inset = input.name.startsWith('producer/') ? 90000 / 12700 / .75 : 9.6;
    result.nativeStyles.insets.forEach((value, index) => expect(value).toBeCloseTo(index < 2 ? inset : inset / 2, 3));
    expect(result.nativeStyles.margin).toBe('0px'); expect(result.nativeStyles.fonts).toBeGreaterThan(0); expect(result.nativeStyles.defaultSize).toBe('24px');
    const [none, normal, shape] = result.cases;
    expect(none).toMatchObject({ width: 100, height: 36, scale: 1, overflow: true });
    expect(normal).toMatchObject({ width: 100, height: 36, overflow: false }); expect(normal!.scale).toBeLessThan(1);
    expect(shape!.height).toBeGreaterThan(36); expect(shape!.geometryHeight).toBeCloseTo(shape!.height, 5); expect(shape!.overflow).toBe(false);
    for (const entry of result.cases) {
      [Math.sqrt(3) / 2, .5, -.5, Math.sqrt(3) / 2].forEach((value, index) => expect(entry.matrix[index]).toBeCloseTo(value, 5));
      expect(entry.text).toBe('中文与 Latin 123，自动调整文字框。'.repeat(5)); expect(entry.nodes).toBe(entry.actualNodes);
    }
    expect(result.leading[1]!.font).toBeCloseTo(result.leading[0]!.font * .5, 3);
    expect(result.leading[1]!.height).toBeCloseTo(result.leading[0]!.height * .375, 0);
    expect(result.listValues.markers).toEqual([{ text: '3.', left: '-10px', position: 'absolute' }, { text: '4.', left: '-10px', position: 'absolute' }]);
    expect(result.listValues.sourceText).toBe('OneTwo'); expect(result.listValues.nodes).toBe(result.listValues.actualNodes);
    expect(result.insetOverflow).toEqual({ width: 1, overflow: true });
    expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded'] });
  }
});

test('explicit native drawing text follows independent fixed-font line advances and empty paragraph sizes', async ({ page }) => {
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureNativeDrawingText(snapshot), nativeText);
  expect(actual).toHaveLength(textEvidence.cases.length);
  for (const [index, entry] of actual.entries()) {
    const reference = textEvidence.cases[index]!;
    expect(entry.unchanged).toBe(true); expect(entry.sourceText).toBe(reference.paragraphs.join('')); expect(entry.placeholders).toBe(reference.paragraphs.flatMap(text => text.split('\n')).filter(text => !text).length);
    expect(entry.fallback).toBe(false);
    expect(entry.nodes).toBe(entry.declared); expect(entry.exact).toBe(entry.textNodes); expect(entry.limited).toBe('ResourceLimitExceeded');
    // Producer-only default font/defPPr differences are recorded separately.
    if (reference.fontPt === null && !['level-22', 'both'].includes(reference.listStyle) || reference.paragraphs.includes('') && reference.endPt === null) continue;
    expect(Math.abs(entry.lineAdvancePt - reference.observed.lineAdvancePt), `case ${index}`).toBeLessThan(.75);
    entry.words.forEach((word, wordIndex) => {
      const expected = reference.observed.words[wordIndex]!; expect(word.text).toBe(expected.text);
      word.boundsPt.forEach((value, axis) => expect(Math.abs(value - expected.boundsPt[axis]!), `case ${index}, ${word.text}, axis ${axis}`).toBeLessThan(1.5));
    });
  }
});

test('drawing text fallback accounts for transient probes and empty lines without inventing source offsets', async ({ page }) => {
  const result = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingTextFallback(snapshot), inputs.find(input => input.name === '08-images-shapes')!.snapshot);
  expect(result).toMatchObject({ error: 'ResourceLimitExceeded', fallback: true, unchanged: true, sourceText: 'Alpha\nBeta', placeholders: 2, cleaned: true });
  expect(result.nodes).toBe(result.declared);
  expect(result.paragraphs[1]).toBeGreaterThan(0);
  expect(result.paragraphs[0]).toBeCloseTo(result.paragraphs[1]! * 2, 1);
  expect(result.paragraphs[2]).toBeCloseTo(result.paragraphs[1]!, 1);
});

test('bounded auxiliary projections preserve native drawing and comment text without activating image paths', async ({ page }) => {
  const requests: string[] = []; page.on('request', request => requests.push(request.url()));
  for (const input of inputs.filter(input => !input.name.startsWith('line-'))) {
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
