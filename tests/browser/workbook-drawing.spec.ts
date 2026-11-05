import { expect, test } from '@playwright/test';
import { readFileSync } from 'node:fs';
import { execFileSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { fileURLToPath } from 'node:url';
import { build } from 'vite';
import { importXlsx } from "@dagent-office/office";
import { openDocument, createWorkbookNumberFormatter } from "@dagent-office/core";
import type { WorkbookSnapshot } from '@dagent-office/core';
import { readXlsxRecords } from '../../packages/office/src/xlsx/read.js';
import { readDocumentZip, writeOpcZip } from '../../packages/office/src/document-zip.js';
import { projectXlsxRecords } from '../../packages/office/tests/helpers/xlsx-record-model.js';
import type { measureContentSource, rejectContentSource } from './fixtures/workbook-content-source.js';
import type { measureImageView, disposeImageView } from './fixtures/workbook-image-view.js';
import type { measureObjectGeometry } from './fixtures/workbook-object-geometry.js';
import type { measureDrawingGradients } from './fixtures/drawing-gradient.js';
import type { measureDrawingText, measureDrawingTextLayout, measureNativeDrawingText, measureDrawingTextFallback, measureDrawingTextOverflow, measureDrawingOverflowEdges } from './fixtures/workbook-drawing-text.js';
import type { measureWorkbookSurface, mountWorkbookSurfacePixels, disposeWorkbookSurfacePixels } from './fixtures/workbook-surface.js';
import type { measureWorkbookShadows, mountWorkbookShadowPixels, disposeWorkbookShadowPixels } from './fixtures/workbook-shadow.js';
import type { measureWorkbookCharts, measureWorkbookChartTitles, measureWorkbookChartPositions, measureWorkbookChartTitleContents, measureWorkbookComposedChart, measureWorkbookChartFormats } from './fixtures/workbook-chart.js';
import { bindWorkbookChartData } from "../../packages/render/src/workbook/chart-values.js";
import type { BasicChartData } from '../../packages/render/src/chart.js';
import { modelChartSource, DEFAULT_CHART_READ_LIMITS } from '../../packages/render/src/workbook/chart-source.js';

declare global { interface Window { workbookChartHost: { measureWorkbookCharts: typeof measureWorkbookCharts; measureWorkbookChartTitles: typeof measureWorkbookChartTitles; measureWorkbookChartTitleContents: typeof measureWorkbookChartTitleContents; measureWorkbookComposedChart: typeof measureWorkbookComposedChart; measureWorkbookChartFormats: typeof measureWorkbookChartFormats; measureWorkbookChartPositions: typeof measureWorkbookChartPositions } } }

declare global { interface Window { workbookShadowHost: { measureWorkbookShadows: typeof measureWorkbookShadows; mountWorkbookShadowPixels: typeof mountWorkbookShadowPixels; disposeWorkbookShadowPixels: typeof disposeWorkbookShadowPixels } } }

declare global { interface Window { workbookContentHost: { measureContentSource: typeof measureContentSource; rejectContentSource: typeof rejectContentSource; measureImageView: typeof measureImageView; disposeImageView: typeof disposeImageView; measureObjectGeometry: typeof measureObjectGeometry; measureDrawingGradients: typeof measureDrawingGradients; measureDrawingText: typeof measureDrawingText; measureDrawingTextLayout: typeof measureDrawingTextLayout; measureNativeDrawingText: typeof measureNativeDrawingText; measureDrawingTextFallback: typeof measureDrawingTextFallback; measureDrawingTextOverflow: typeof measureDrawingTextOverflow; measureDrawingOverflowEdges: typeof measureDrawingOverflowEdges; measureWorkbookSurface: typeof measureWorkbookSurface; mountWorkbookSurfacePixels: typeof mountWorkbookSurfacePixels; disposeWorkbookSurfacePixels: typeof disposeWorkbookSurfacePixels } } }
const inputs: { name: string; snapshot: WorkbookSnapshot; text: Record<string, string>; resources: Record<string, string> }[] = [];
const chartPlots = new Map<string, BasicChartData[]>();
let referenceTitles: { snapshot: WorkbookSnapshot; bindings: ReturnType<typeof bindWorkbookChartData> };
let formatCharts: typeof referenceTitles;
let code = '';
let nativeText: WorkbookSnapshot;
let overflowText: WorkbookSnapshot;
let anchoredText: WorkbookSnapshot;
const anchoringEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-text-anchoring.json', 'utf8')) as { sourceSha256: string; cases: { paragraphs: string[]; anchorCenter: boolean; observed: { words: { text: string; boundsPt: number[] }[] } }[] };
let nativeShadows: WorkbookSnapshot, nativeBlur: WorkbookSnapshot;
const textEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-drawing-text-semantics.json', 'utf8')) as { sourceSha256: string; cases: { fontPt: number | null; endPt: number | null; paragraphs: string[]; listStyle: string; observed: { lineAdvancePt: number; words: { text: string; boundsPt: number[] }[] } }[] };
const shadowEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-drawing-shadow-semantics.json', 'utf8')) as { sourceSha256: string; cases: { rotation: number; flipH: boolean; flipV: boolean; rotateWithShape: boolean; direction: number; distancePt: number; observed: { offsetPt: number[] } }[] };
const blurEvidence = JSON.parse(readFileSync('docs/evidence/xlsx-shadow-blur-semantics.json', 'utf8')) as { sourceSha256: string; cases: { blurPt: number; observed: { sharpEdgeOffsetPt?: number; samples: { offsetPt: number; rgb: number[] }[] } }[] };
test.beforeAll(async () => {
  overflowText = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-text-overflow.xlsx'), { documentId: 'text-overflow' }), { kind: 'workbook' }).snapshot();
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-content-source.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookDrawingProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw Error('Unexpected drawing fixture bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output; code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
  const nativeBytes = execFileSync('python3', ['-c', "import importlib.util,sys; s=importlib.util.spec_from_file_location('drawing_text_inputs','scripts/check-xlsx-drawing-text.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))"], { maxBuffer: 16 * 1024 * 1024 });
  expect(createHash('sha256').update(nativeBytes).digest('hex')).toBe(textEvidence.sourceSha256);
  nativeText = projectXlsxRecords(await readXlsxRecords(nativeBytes)).model.snapshot();
  const anchoringBytes = readFileSync('packages/office/tests/fixtures/xlsx-text-anchoring.xlsx');
  expect(createHash('sha256').update(anchoringBytes).digest('hex')).toBe(anchoringEvidence.sourceSha256);
  anchoredText = openDocument(await importXlsx(anchoringBytes, { documentId: 'native-text-anchoring' }), { kind: 'workbook' }).snapshot();
  for (const [name, evidence] of [['drawing-effects', shadowEvidence], ['shadow-blur', blurEvidence]] as const) {
    const bytes = execFileSync('python3', ['-c', `import importlib.util,sys; s=importlib.util.spec_from_file_location('shadow_inputs','scripts/check-xlsx-${name}.py'); m=importlib.util.module_from_spec(s); s.loader.exec_module(m); sys.stdout.buffer.write(m.generate(m.inputs()))`], { maxBuffer: 16 * 1024 * 1024 });
    expect(createHash('sha256').update(bytes).digest('hex')).toBe(evidence.sourceSha256);
    const snapshot = projectXlsxRecords(await readXlsxRecords(bytes)).model.snapshot(); if (name === 'drawing-effects') nativeShadows = snapshot; else nativeBlur = snapshot;
  }
  for (const prefix of ['', 'producer/']) for (const name of ['07-three-charts', '08-images-shapes', '09-classic-notes', '10-modern-threads']) {
    const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync(`examples/xlsx/${prefix}${name}.xlsx`))), snapshot = model.snapshot();
    const ids = [...(snapshot.metadata.workbook.stories ?? []).map(story => story.rootNodeId), ...(snapshot.metadata.workbook.objects ?? []).map(object => object.nodeId)];
    inputs.push({ name: prefix + name, snapshot, text: Object.fromEntries(ids.map(id => [id, model.getNode(id)!.text])), resources: Object.fromEntries(snapshot.manifest.assets.map(asset => [asset.path, Buffer.from(model.getAsset(asset.path)!).toString('base64')])) });
    if (name === '07-three-charts') {
      const definitions = snapshot.metadata.workbook.charts!, response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: definitions.map(chart => chart.chartId) }, new AbortController().signal);
      const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
      chartPlots.set(prefix + name, bindWorkbookChartData(snapshot, response, formatter).map(chart => {
        if (chart.kind !== 'available') throw Error('Independent chart source must be available'); return chart.data;
      })); formatter.dispose();
    }
  }
  {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/07-three-charts.xlsx'))), part = 'xl/charts/chart1.xml', encoder = new TextEncoder();
    const title = (text: string, scale: number) => `<c:title><c:tx><c:rich><a:bodyPr><a:normAutofit fontScale="${scale}"/></a:bodyPr><a:lstStyle/><a:p><a:r><a:rPr sz="1400"/><a:t>${text}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
    const original = new TextDecoder().decode(files.get(part)!);
    const xml = original.replace('<c:axPos val="b"/>', '<c:axPos val="b"/>' + title('季度😀', 75000)).replace('<c:axPos val="l"/>', '<c:axPos val="l"/>' + title('收入禰󠄀', 50000));
    expect(xml).not.toBe(original); files.set(part, encoder.encode(xml));
    const snapshot = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files))).model.snapshot();
    const text = Object.fromEntries(snapshot.metadata.workbook.charts!.map((chart, index) => [chart.nodeId, '图表 ' + chart.kind + (index === 0 ? '季度😀收入禰󠄀' : '')]));
    inputs.push({ name: 'chart-title-layout', snapshot, text, resources: {} });
  }
  {
    const files = new Map(readDocumentZip(readFileSync('examples/xlsx/07-three-charts.xlsx'))), part = 'xl/charts/chart1.xml', encoder = new TextEncoder(), decoder = new TextDecoder();
    const title = (cell: string) => `<c:title><c:tx><c:strRef><c:f>Data!$${cell}$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>stale cache</c:v></c:pt></c:strCache></c:strRef></c:tx><c:txPr><a:bodyPr lIns="0" rIns="0" tIns="0" bIns="0"><a:normAutofit fontScale="75000"/></a:bodyPr><a:lstStyle/><a:p><a:pPr algn="ctr"><a:defRPr sz="1600"/></a:pPr></a:p></c:txPr></c:title>`;
    files.set(part, encoder.encode(decoder.decode(files.get(part)!).replace(/<c:title>.*?<\/c:title>/, title('A')).replace('<c:axPos val="b"/>', '<c:axPos val="b"/>' + title('B')).replace('<c:axPos val="l"/>', '<c:axPos val="l"/>' + title('C'))));
    const sheet = 'xl/worksheets/sheet1.xml'; files.set(sheet, encoder.encode(decoder.decode(files.get(sheet)!).replace('季度</t>', '当前😀禰󠄀\n标题</t>')));
    const { model } = projectXlsxRecords(await readXlsxRecords(writeOpcZip(files))), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    referenceTitles = { snapshot, bindings: bindWorkbookChartData(snapshot, response, formatter) }; formatter.dispose();
  }
  {
    const model = openDocument(await importXlsx(readFileSync('packages/office/tests/fixtures/xlsx-chart-formats.xlsx'), { documentId: 'chart-formats' }), { kind: 'workbook' }), snapshot = model.snapshot();
    const response = await modelChartSource(model).read({ ...DEFAULT_CHART_READ_LIMITS, documentId: snapshot.manifest.documentId, revision: snapshot.manifest.revision, chartIds: snapshot.metadata.workbook.charts!.map(chart => chart.chartId) }, new AbortController().signal);
    const formatter = createWorkbookNumberFormatter({ dateSystem: snapshot.metadata.workbook.dateSystem });
    formatCharts = { snapshot, bindings: bindWorkbookChartData(snapshot, response, formatter) }; formatter.dispose();
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

test('workbook chart graphics use current values and isolated source styles with bounded failure cleanup', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('07-three-charts'))) {
    const results = await page.evaluate(({ snapshot, plots }) => window.workbookChartHost.measureWorkbookCharts(snapshot, plots), { snapshot: input.snapshot, plots: chartPlots.get(input.name)! });
    expect(results.map(result => result.kind)).toEqual(['column', 'line', 'pie']);
    for (const result of results) {
      expect(result.reported).toBe(result.nodes); expect(result.peak).toBeGreaterThan(result.nodes); expect(result.exact).toBe(true); expect(result.unchanged).toBe(true); expect(result.sourceClasses).toBe(0);
      expect(result.svgSize).toEqual([480, 293]); expect(result.fonts.every(value => Math.abs(parseFloat(value) - 13) < .02), JSON.stringify(result.fonts)).toBe(true);
      expect(result.colors.every(value => value.replace(/\s/g, '') === 'rgb(12,34,56)'), JSON.stringify(result.colors)).toBe(true);
      expect(result.errors).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'format rejected', 'InvalidOperation', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded']);
      expect(result.negativePies).toHaveLength(result.kind === 'pie' ? 6 : 0);
      for (const pie of result.negativePies) {
        expect(pie.geometry).toEqual(pie.absolute); expect(pie.labels).toEqual(pie.values.filter(value => value !== null && value !== 0));
        expect(pie.fractions.every(value => value > 0 && value <= 1)).toBe(true);
        expect(pie.fractions.reduce((a, b) => a + b, 0)).toBeCloseTo(pie.labels.length ? 1 : 0, 12);
        const negativePoints = pie.values.filter(value => value !== null && value < 0).length;
        expect(pie.diagnostics).toEqual(negativePoints ? [{ code: 'negative-pie-magnitude', chartId: expect.any(String), negativePoints }] : []);
        expect(pie.positiveDiagnostics).toEqual([]); expect(pie.immutable).toBe(true);
      }
      for (const fraction of result.smallScales) expect(fraction).toBeGreaterThan(.7);
      expect(result.zeroScales).toEqual(result.kind === 'column' ? [0, 0] : []);
      expect(result.titleProjection.map(part => part.text).join('')).toBe(result.titleText);
      for (const part of result.titleProjection) { expect(part.owned).toBe(true); expect(part.to - part.from).toBe(part.text.length); expect(result.titleText!.slice(part.from, part.to)).toBe(part.text); }
      expect(result.text).toEqual(expect.arrayContaining(['第一季度', '第二季度', '第三季度']));
    }
    expect(results[0]!.bars).toHaveLength(6); expect(results[0]!.bars[2]! / results[0]!.bars[0]!).toBeCloseTo(4, 4);
  }
});

test('workbook rich chart titles share source-aware text layout, fit budgets and failure isolation', async ({ page }) => {
  for (const input of inputs.filter(input => input.name.endsWith('07-three-charts') || input.name === 'chart-title-layout')) {
    const result = await page.evaluate(snapshot => window.workbookChartHost.measureWorkbookChartTitles(snapshot), input.snapshot);
    expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'InvalidOperation', 'InvalidDocument', 'InvalidOperation', 'ResourceLimitExceeded'] });
    expect(result.native.length).toBeGreaterThanOrEqual(3);
    for (const title of result.native) {
      expect(title.unchanged && title.owned).toBe(true); expect(title.nodes).toBe(title.declared); expect(title.fonts).toBeGreaterThan(0);
      expect(title.sourceId).toBe(title.expectedId); expect(title.widthPt).toBe(240); expect(title.units).toEqual(title.lengths);
      expect(title.reflected).toBe('matrix(1, 0, 0, 1, 0, 0)');
    }
    if (input.name === 'chart-title-layout') {
      expect(result.native.filter(title => title.role !== 'title').map(title => ({ role: title.role, text: title.text, scale: title.scale }))).toEqual([
        { role: 'category-axis-title', text: '季度😀', scale: .75 }, { role: 'value-axis-title', text: '收入禰󠄀', scale: .5 },
      ]);
    }
    expect(result.fits.every(item => item.unchanged && item.sourceText === 'Long title with 中文😀禰󠄀 and several words that need multiple lines.')).toBe(true);
    expect(result.fits[0]!.overflow).toBe(true); expect(result.fits[0]!.scale).toBe(1);
    expect(result.fits[1]!.scale).toBeLessThan(1); expect(result.fits[1]!.width).toBe(120); expect(result.fits[1]!.height).toBe(30);
    expect(result.fits[2]!.height).toBeGreaterThan(30); expect(result.fits[2]!.scale).toBe(1);
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
    expect(result.zeroList).toEqual({ margin: '0px', indent: '0px', markerLeft: '0px', markerWidth: '0px', sourceText: 'OneTwo' });
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

test('native text overflow clips actual ink independently by axis and keeps complete source offsets behind derived ellipses', async ({ page }, info) => {
  const reference = JSON.parse(readFileSync('docs/evidence/xlsx-text-overflow.json', 'utf8')) as { cases: { horizontal: string; vertical: string; anchor: string; paragraphs: string[] }[] };
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingTextOverflow(snapshot), overflowText);
  expect(actual).toMatchObject({ invalid: 'InvalidDocument', limited: 'ResourceLimitExceeded', geometryLimited: 'ResourceLimitExceeded', cleaned: true });
  expect(actual.results).toHaveLength(24);
  const image = await page.locator('#workbook-overflow-probe').screenshot({ path: info.outputPath('overflow-ink.png') });
  await info.attach('native-text-overflow', { body: image, contentType: 'image/png' });
  const ink = await page.evaluate(async bytes => {
    const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
    try {
      image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
      const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, image.width, image.height).data;
      return Array.from({ length: 24 }, (_, index) => {
        const x0 = index % 4 * 300, y0 = Math.floor(index / 4) * 180;
        let left = 300, top = 180, right = -1, bottom = -1, count = 0;
        for (let y = 0; y < 180; y++) for (let x = 0; x < 300; x++) {
          const offset = ((y0 + y) * image.width + x0 + x) * 4;
          if (Math.max(pixels[offset]!, pixels[offset + 1]!, pixels[offset + 2]!) < 180) {
            left = Math.min(left, x); right = Math.max(right, x); top = Math.min(top, y); bottom = Math.max(bottom, y); count++;
          }
        }
        return [left, top, right, bottom, count];
      });
    } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
  }, [...image]);
  for (const [index, result] of actual.results.entries()) {
    const expected = reference.cases[index]!, bounds = ink[index]!;
    expect(result.unchanged).toBe(true); expect(result.sourceText).toBe(expected.paragraphs.join(''));
    expect(result.nodes).toBe(result.declared); expect(result.markerSourceIds).toBe(false); expect(result.overflow).toBe(true);
    expect(result.markerCount).toBe(expected.vertical === 'ellipsis' ? 1 : 0);
    if (expected.vertical === 'ellipsis' && expected.anchor === 't') {
      expect(result.markerBounds[0]!.top).toBeGreaterThan(12); expect(result.markerBounds[0]!.top).toBeLessThan(34);
      expect(result.markerBounds[0]!.left).toBeGreaterThan(20);
    }
    for (const [from, to, length] of result.offsets) expect(to! - from!).toBe(length);
    if (expected.horizontal === 'clip') { expect(bounds[0]).toBeGreaterThanOrEqual(40); expect(bounds[2]).toBeLessThan(201); expect(result.paintBounds.width).toBe(120); }
    if (expected.vertical !== 'overflow') { expect(bounds[1]).toBeGreaterThanOrEqual(60); expect(bounds[3]).toBeLessThan(114); expect(result.paintBounds).toMatchObject({ y: 0, height: 40 }); }
    if (expected.vertical === 'overflow') expect(expected.anchor === 'b' ? bounds[1]! < 60 : bounds[3]! >= 114).toBe(true);
  }
  // A real derived glyph and its reserved gap must change visible ink.
  for (const start of [0, 12]) for (let offset = 0; offset < 4; offset++) expect(ink[start + 8 + offset]).not.toEqual(ink[start + 4 + offset]);
});

test('native drawing block centering preserves paragraph alignment and vertical anchoring', async ({ page }) => {
  const actual = await page.evaluate(snapshot => window.workbookContentHost.measureNativeDrawingText(snapshot), anchoredText);
  expect(actual).toHaveLength(anchoringEvidence.cases.length);
  for (const [index, item] of actual.entries()) {
    const expected = anchoringEvidence.cases[index]!;
    expect(item).toMatchObject({ unchanged: true, sourceText: expected.paragraphs.join(''), limited: 'ResourceLimitExceeded' });
    expect(item.nodes).toBe(item.declared); expect(item.exact).toBe(item.textNodes);
    for (const [wordIndex, word] of item.words.entries()) {
      const reference = expected.observed.words[wordIndex]!; expect(word.text).toBe(reference.text);
      for (const [axis, value] of word.boundsPt.entries()) expect(Math.abs(value - reference.boundsPt[axis]!), `case ${index + 1}, ${word.text}, axis ${axis}`).toBeLessThan(1.5);
    }
    if (expected.anchorCenter) {
      const left = Math.min(...item.words.map(word => word.boundsPt[0]!)), right = Math.max(...item.words.map(word => word.boundsPt[2]!));
      expect(Math.abs((left + right) / 2 - 60)).toBeLessThan(.5);
    }
  }
});

for (const invalidCanvasShorthand of [false, true]) test(`drawing text fallback accounts for transient probes and empty lines with ${invalidCanvasShorthand ? 'valid CSS outside Canvas shorthand' : 'missing Canvas font bounds'}`, async ({ page }) => {
  const result = await page.evaluate(({ snapshot, invalidCanvasShorthand }) => window.workbookContentHost.measureDrawingTextFallback(snapshot, invalidCanvasShorthand), { snapshot: inputs.find(input => input.name === '08-images-shapes')!.snapshot, invalidCanvasShorthand });
  expect(result).toMatchObject({ error: 'ResourceLimitExceeded', fallback: true, unchanged: true, sourceText: 'Alpha\nBeta', placeholders: 2, cleaned: true });
  expect(result.nodes).toBe(result.declared);
  expect(result.paragraphs[1]).toBeGreaterThan(0);
  expect(result.paragraphs[0]).toBeCloseTo(result.paragraphs[1]! * 2, 1);
  expect(result.paragraphs[2]).toBeCloseTo(result.paragraphs[1]!, 1);
});

test('workbook surfaces isolate source important rules, host styles and concurrent candidate lifecycles', async ({ page }) => {
  const result = await page.evaluate(snapshot => window.workbookContentHost.measureWorkbookSurface(snapshot), inputs.find(input => input.name === '08-images-shapes')!.snapshot);
  expect(result).toMatchObject({ errors: ['ResourceLimitExceeded', 'InvalidOperation', 'OperationCancelled', 'OperationCancelled'], kept: true, independent: true, released: true });
  expect(result.values.map(value => value.sourceColor)).toEqual(['rgb(13, 75, 91)', 'rgb(9, 33, 77)']);
  for (const value of result.values) {
    expect(value).toMatchObject({ frame: [400, 200], view: [400, 200], object: [160, 120], geometry: [160, 120], sourceWidth: '1332px', paragraphFont: '0px', border: '0px', hidden: 'hidden', connectedAfterReadyAbort: true, unchanged: true });
    expect(Math.abs(parseFloat(value.sourceFont) - 14 / .75)).toBeLessThan(.02); expect(value.paragraphWidth).toBeLessThan(160);
    expect(value.selection).toMatchObject({ text: value.selection.expected, range: true, errors: ['InvalidOperation', 'InvalidOperation'], unchanged: true });
  }
});

test('isolated workbook surfaces paint local SVG gradient references in actual page pixels', async ({ page }) => {
  await page.evaluate(snapshot => window.workbookContentHost.mountWorkbookSurfacePixels(snapshot), inputs.find(input => input.name === '08-images-shapes')!.snapshot);
  try {
    const png = await page.locator('#workbook-surface-pixel-probe iframe').screenshot();
    const pixels = await page.evaluate(async bytes => {
      const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
      try {
        image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
        const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
        return { width: image.width, height: image.height, values: [40, 200, 360].map(x => [...context.getImageData(x, 100, 1, 1).data]) };
      } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
    }, [...png]);
    expect(pixels).toMatchObject({ width: 400, height: 200 });
    for (const [index, x] of [40, 200, 360].entries()) {
      const expected = [255 * (1 - (x + .5) / 400), 0, 255 * (x + .5) / 400, 255];
      pixels.values[index]!.forEach((value, channel) => expect(Math.abs(value - expected[channel]!)).toBeLessThan(2));
    }
  } finally { await page.evaluate(() => window.workbookContentHost.disposeWorkbookSurfacePixels()); }
});

test('worksheet shadows preserve page offsets through native rotations and reflections with bounded raster allocation', async ({ page }) => {
  const result = await page.evaluate(snapshot => window.workbookShadowHost.measureWorkbookShadows(snapshot), nativeShadows);
  expect(result.errors.slice(0, 5)).toEqual(Array(5).fill('InvalidDocument'));
  // Old Chromium saturates enormous CSS numbers before computed-style access;
  // it must still fail admission instead of allocating an enormous filter.
  expect(['InvalidDocument', 'ResourceLimitExceeded']).toContain(result.errors[5]);
  expect(result.errors.slice(6)).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'UnsupportedFeature']); expect(result.restored).toBe(true);
  expect(result.none).toMatchObject({ filter: 'none', outsetPt: 0, rasterPixels: 0, approximateBlur: false });
  expect(result.cases).toHaveLength(shadowEvidence.cases.length);
  result.cases.forEach((actual, index) => {
    const source = shadowEvidence.cases[index]!;
    expect(actual).toMatchObject({ outset: 15, approximate: false, unchanged: true, filterEnabled: true });
    expect(actual.pixels).toBeGreaterThan(0); expect(actual.pixels).toBeLessThan(16_777_216);
    if (!source.rotateWithShape) actual.worldOffsetPt.forEach((value, axis) => expect(Math.abs(value - source.observed.offsetPt[axis]!)).toBeLessThan(.02));
    else {
      // The source flag retains object-space direction. Calc's contrary
      // producer behavior is recorded, not copied into these expected values.
      const x = source.direction === 0 ? (source.flipH ? -15 : 15) : 0, y = source.direction === 90 ? (source.flipV ? -15 : 15) : 0;
      const expected = source.rotation === 0 ? [x, y] : [Math.sqrt(3) / 2 * x - .5 * y, .5 * x + Math.sqrt(3) / 2 * y];
      actual.worldOffsetPt.forEach((value, axis) => expect(value).toBeCloseTo(expected[axis]!, 4));
    }
  });
});

test('worksheet Gaussian shadow preview reports approximation and checks native edge profiles in actual pixels', async ({ page }) => {
  let browserEdgePt = 0;
  const nativeEdgePt = blurEvidence.cases[0]!.observed.sharpEdgeOffsetPt!;
  for (const [index, reference] of blurEvidence.cases.entries()) {
    const result = await page.evaluate(({ snapshot, index }) => window.workbookShadowHost.mountWorkbookShadowPixels(snapshot, index), { snapshot: nativeBlur, index });
    try {
      const png = await page.locator('#workbook-shadow-pixel-probe iframe').screenshot();
      const raster = await page.evaluate(async ({ bytes, offsets }) => {
        const url = URL.createObjectURL(new Blob([new Uint8Array(bytes)], { type: 'image/png' })), image = new Image();
        try {
          image.src = url; await image.decode(); const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height;
          const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
          const row = Math.round(70 / .75 - .5), values = offsets.map(offset => [...context.getImageData(Math.round((220 + offset) / .75 - .5), row, 1, 1).data].slice(0, 3));
          let edge = NaN;
          for (let x = Math.floor(220 / .75) - 6; x < Math.floor(220 / .75) + 6; x++) {
            const a = context.getImageData(x, row, 1, 1).data[1]!, b = context.getImageData(x + 1, row, 1, 1).data[1]!;
            if (a < 127.5 && b >= 127.5) { edge = (x + .5 + (127.5 - a) / (b - a)) * .75 - 220; break; }
          }
          return { values, edge };
        } finally { image.removeAttribute('src'); URL.revokeObjectURL(url); }
      // Compare the blur kernels after independently measuring each raster's
      // sharp edge. Keep the sub-point placement error separately bounded;
      // do not shift the renderer to reproduce a producer's pixel-grid snap.
      }, { bytes: [...png], offsets: reference.observed.samples.map(sample => sample.offsetPt + (index ? browserEdgePt - nativeEdgePt : 0)) });
      if (!index) { expect(Math.abs(raster.edge)).toBeLessThan(1); browserEdgePt = raster.edge; }
      expect(result.approximate).toBe(reference.blurPt > 0);
      const errors = raster.values.flatMap((pixel, sampleIndex) => {
        const sample = reference.observed.samples[sampleIndex]!;
        // Sharp vector edges have a producer/pixel-grid quantization rim.
        if (!reference.blurPt && Math.abs(sample.offsetPt) < 1) return [];
        return pixel.map((value, channel) => Math.abs(value - sample.rgb[channel]!));
      });
      expect(Math.max(...errors)).toBeLessThanOrEqual(24);
    } finally { await page.evaluate(() => window.workbookShadowHost.disposeWorkbookShadowPixels()); }
  }
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

test('workbook chart compaction preserves original point styles and distinguishes errors from skipped points', async ({ page }) => {
  const input = inputs.find(input => input.name === '07-three-charts')!;
  const result = await page.evaluate(snapshot => window.workbookChartHost.measureWorkbookChartPositions(snapshot), input.snapshot);
  expect(result.maskedColors).toEqual(['rgba(237,125,49,1)', 'rgba(237,125,49,1)']);
  expect(result.colors).toEqual(['rgb(0, 0, 255)', 'rgb(255, 0, 0)']);
  expect(result.plain.match(/M/g)).toHaveLength(1); expect(result.plain.match(/L/g)).toHaveLength(4);
  expect(result.missing.match(/M/g)).toHaveLength(2); expect(result.missing.match(/L/g)).toHaveLength(1);
  expect(result.hiddenPaths).toBe(0); expect(result.hiddenDiagnostics).toEqual([]);
  expect(result.errors).toEqual(['InvalidOperation', 'InvalidOperation', 'InvalidOperation']); expect(result.unchanged).toBe(true);
});


test('current reference chart titles share measured typography without cached text offsets', async ({ page }) => {
  const result = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookChartTitleContents(snapshot, bindings), referenceTitles);
  expect(result).toMatchObject({ unchanged: true, cleaned: true, empty: { projections: 0, visible: '' }, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidOperation', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded'] });
  expect(result.fonts).toBeGreaterThan(0);
  expect(result.results.slice(0, 3).map(item => item.text)).toEqual(['当前😀禰󠄀\n标题', '交付', '计划']);
  for (const item of result.results) {
    expect(item.nodes).toBe(item.declared); expect(item.sourceClasses).toBe(0); expect(item.sourceId).toBe(item.expectedId);
    if (item.reference) { expect(item.projections).toHaveLength(0); expect(item.scale).toBe(.75); expect(item.fontSizes.every(size => Math.abs(parseFloat(size) - 16) < .02)).toBe(true); }
    else expect(item.projections.map(projection => projection.text).join('')).toBe(item.text);
  }
});


test('composed workbook charts reserve separate title, axis and legend boxes and publish atomically', async ({ page }) => {
  const result = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookComposedChart(snapshot, bindings), referenceTitles);
  expect(result).toMatchObject({ unchanged: true, cleaned: true, errors: ['ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'ResourceLimitExceeded', 'InvalidOperation', 'format rejected'] });
  const overlap = (a: { x: number; y: number; width: number; height: number }, b: typeof a) => Math.min(a.x + a.width, b.x + b.width) - Math.max(a.x, b.x) > .1 && Math.min(a.y + a.height, b.y + b.height) - Math.max(a.y, b.y) > .1;
  for (const item of result.results) {
    expect(item.nodes).toBe(item.declared); expect(item.styles).toBe(0); expect(item.unchanged).toBe(true);
    for (const key of ['x', 'y', 'width', 'height'] as const) expect(item.plot[key]).toBeCloseTo(item.layout.plot[key], 1);
    expect(item.titles.map(title => title.text)).toEqual(['当前😀禰󠄀\n标题', '交付', '计划']);
    for (const title of item.titles) {
      expect(title.projections).toBe(0); expect(overlap(title.rect, item.plot)).toBe(false);
      expect(title.rect.x).toBeGreaterThanOrEqual(-.1); expect(title.rect.y).toBeGreaterThanOrEqual(-.1);
      expect(title.rect.x + title.rect.width).toBeLessThanOrEqual(420.1); expect(title.rect.y + title.rect.height).toBeLessThanOrEqual(300.1);
      if (item.layout.legend) expect(overlap(title.rect, item.layout.legend)).toBe(false);
    }
    const category = item.titles.find(title => title.role === 'category-axis-title')!.rect, value = item.titles.find(title => title.role === 'value-axis-title')!.rect;
    expect(category.y < item.plot.y).toBe(item.opposite); expect(value.x > item.plot.x).toBe(item.opposite);
  }
});


test('chart axes and labels resolve source-linked formats from current cells and independent native formats', async ({ page }) => {
  const results = await page.evaluate(({ snapshot, bindings }) => window.workbookChartHost.measureWorkbookChartFormats(snapshot, bindings), formatCharts);
  expect(results).toHaveLength(20); expect(results.every(result => result.unchanged)).toBe(true);
  for (let mode = 0; mode < 4; mode++) {
    expect(results[mode]!.valueAxis).toEqual(mode === 0 ? ['0.00', '0.25', '0.50', '0.75', '1.00'] : ['0.0%', '25.0%', '50.0%', '75.0%', '100.0%']);
    expect(results[4 + mode]!.categoryAxis).toEqual(mode === 0 ? ['0.10', '0.20', '0.30'] : ['10.0%', '20.0%', '30.0%']);
    expect(results[8 + mode]!.labels).toEqual(mode === 0 ? ['0.25', '0.50', '0.75'] : ['25.0%', '50.0%', '75.0%']);
    expect(results[12 + mode]!.labels).toEqual(mode === 0 ? ['0.17', '0.33', '0.50'] : ['17%', '33%', '50%']);
  }
  for (const result of results.slice(16, 19)) expect(result.valueAxis).toEqual(['0', '0.25', '0.5', '0.75', '1']);
  expect(results[18]!.sourceIndexes).toEqual([1, 2]);
  expect(results[19]!.labels).toEqual(['0.25', '50.0%', '0.75']);
});


test('text overflow preserves cross-span graphemes, RTL and reflections even when no full line fits', async ({ page }) => {
  const results = await page.evaluate(snapshot => window.workbookContentHost.measureDrawingOverflowEdges(snapshot), overflowText);
  expect(results).toHaveLength(24);
  for (const result of results) {
    expect(result, JSON.stringify(result)).toMatchObject({ splitGraphemes: 0, unchanged: true, marker: '…', inside: true });
    expect(result.text).toBe('A禰󠄀é😀WXYZ אבג ABCSecond lineHidden tail'); expect(result.nodes).toBe(result.declared);
  }
});
