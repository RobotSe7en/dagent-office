import { expect, test } from "@playwright/test";
import { build } from "vite";
import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import { readXlsxRecords } from "../../packages/office/src/xlsx/read.js";
import { projectXlsxRecords } from "../../packages/office/tests/helpers/xlsx-record-model.js";
import type { WorkbookViewport } from "../../packages/render/src/workbook/viewport.js";
import type { WorkbookTextFontRun } from "../../packages/render/src/workbook/text-fonts.js";
import type { WorkbookBorderEdge } from "../../packages/render/src/workbook/border-view.js";
import type { WorkbookBorderKind } from "../../packages/render/src/workbook/border-priority.js";
import borderReference from "../../docs/evidence/xlsx-border-semantics.json" with { type: "json" };
import alignmentReference from "../../docs/evidence/xlsx-alignment-semantics.json" with { type: "json" };
import type { TextLayoutCase, textLayoutMeasurement } from "./fixtures/workbook-text-layout.js";
import type { borderSourceMeasurement } from "./fixtures/workbook-border-source.js";
import type { assetMeasurement } from "./fixtures/workbook-assets.js";
import { workbookRasterCases } from "../workbook-image-fixture.js";
import type { measureWorkbookFills } from "./fixtures/workbook-fill.js";
import fillReference from "../../docs/evidence/xlsx-fill-pattern-spec.json" with { type: "json" };
import appearanceReference from "../../docs/evidence/xlsx-cell-appearance.json" with { type: "json" };
import borderSpec from "../../docs/evidence/xlsx-border-spec.json" with { type: "json" };
import type { diagonalMeasurement } from "./fixtures/workbook-diagonals.js";
import rotationReference from "../../docs/evidence/xlsx-rotation-semantics.json" with { type: "json" };
import type { numberTextMeasurement, NumberTextCase } from "./fixtures/workbook-number-text.js";

interface GridHost {
  numberTextMeasurement(cases: readonly NumberTextCase[], locale?: string): ReturnType<typeof numberTextMeasurement>;
  diagonalMeasurement(css: string, classes: readonly string[]): ReturnType<typeof diagonalMeasurement>;
  measureWorkbookFills(css: string, classes: readonly string[]): ReturnType<typeof measureWorkbookFills>;
  assetMeasurement(base64: string, mimeType: string): ReturnType<typeof assetMeasurement>;
  borderSourceMeasurement(): ReturnType<typeof borderSourceMeasurement>;
  textLayoutMeasurement(cases: readonly TextLayoutCase[]): ReturnType<typeof textLayoutMeasurement>;
  render(viewport: WorkbookViewport, rtl?: boolean): Promise<boolean>;
  inspect(): { cellCount: number; borderLines: number; accessibleCells: number; rows: number[]; sourceUnchanged: boolean; merges: { pane?: string; text: string | null; clip: { x: number; y: number; width: number; height: number }; body: { x: number; y: number; width: number; height: number }; hit?: { row: number; column: number } }[] };
  failedPaint(): boolean; dispose(): void;
  numericMeasurement(): Promise<{ rounded: { kind: string; text?: string }; scientific: { kind: string; text?: string }; actual: number; predicted: number; overflow: boolean; fill?: { kind: string; text?: string } }>;
  columnFontMeasurement(): Promise<{ measured: { cssClass: string; maximumDigitWidthPx: number; fonts: { requestedFamily: string; substituted: boolean }[] }; fallback: { cssClass: string; maximumDigitWidthPx: number }; actual: number; unrelatedUnchanged: boolean; retainedProbes: number }>;
  colorMeasurement(): { values: (string | undefined)[]; errors: string[]; fills: string[]; fillErrors: string[]; unchanged: boolean };
  sourceStyleMeasurement(): Promise<{ text: string | null; font: string; size: string; italic: string; background: string; bold: string; foreground: string; plain: string; blockRow?: string; row?: string; bound: boolean; before: string; disposed: boolean; noPartial: boolean; retained: number; rejected: string[] }>;
  textFontMeasurement(): Promise<{ runs: readonly WorkbookTextFontRun[]; rich: readonly WorkbookTextFontRun[]; withBreak: readonly WorkbookTextFontRun[]; limited: boolean; unchanged: boolean; frozen: boolean }>;
  textViewMeasurement(): Promise<{ rich: TextRunView[]; mixed: TextRunView[]; decorated: TextRunView[]; decoration: { strike: string; strikeStyle: string; underline: string; underlineStyle: string; plain: string }; text: string | null; failures: string[]; unchanged: boolean; frozen: boolean }>;
  borderViewMeasurement(edges: readonly WorkbookBorderEdge[]): Promise<{ strokes: { from: number[]; to: number[]; width: number; dash: number[]; cap: string | null }[]; difference: number; maskDifference: number; splitDifference: number; splitMaskDifference: number; painted: number; doubleRows: number[]; failures: string[]; rasterImages: string[]; rasterSizes: number[][] }>;
  textDirectionMeasurement(): { values: string[]; explicit: string[]; wholeStringDirection: string; retainedText: string | null; failures: string[]; remaining: number; ownerDirection: string };
}
interface TextRunView { text: string; from: number; to: number; family: string; weight: string; color: string; x: number; y: number; width: number; height: number }
declare global { interface Window { workbookGridHost: GridHost } }
let code = '';
test.beforeAll(async () => {
  const bundle = await build({ configFile: false, build: { write: false, minify: false, lib: { entry: fileURLToPath(new URL('./fixtures/workbook-grid-host.ts', import.meta.url)), formats: ['iife'], name: 'WorkbookGridProbe' } } });
  if (!Array.isArray(bundle) && !('output' in bundle)) throw new Error('Unexpected workbook grid probe bundle');
  const output = Array.isArray(bundle) ? bundle.flatMap(part => part.output) : bundle.output;
  code = output.filter(part => part.type === 'chunk').map(part => part.code).join('\n');
});
test.beforeEach(async ({ page }) => { await page.goto('/'); await page.addScriptTag({ content: code }); await page.waitForFunction(() => !!window.workbookGridHost); });

for (const rtl of [false, true]) test(`merged frozen fragments preserve full content geometry in ${rtl ? 'RTL' : 'LTR'}`, async ({ page }) => {
  const report = await page.evaluate(async rtl => {
    await window.workbookGridHost.render({ widthPt: 100, heightPt: 80, scrollLeftPt: 40, scrollTopPt: 30 }, rtl);
    return window.workbookGridHost.inspect();
  }, rtl);
  expect(report.merges).toHaveLength(4); expect(report.cellCount).toBe(24); expect(report.accessibleCells).toBe(21);
  expect(report.rows).toEqual([...report.rows].sort((a, b) => a - b)); expect(report.sourceUnchanged).toBe(true);
  for (const cell of report.merges) {
    expect(cell.text).toBe('合并😀禰󠄀'); expect(cell.hit).toEqual({ row: 0, column: 0 });
    expect(cell.body.width).toBeCloseTo(140, 1); expect(cell.body.height).toBeCloseTo(70, 1);
  }
  const body = report.merges.find(cell => cell.pane === 'body')!;
  expect(body.clip.x).toBeCloseTo(rtl ? 0 : 20, 1); expect(body.clip.y).toBeCloseTo(20, 1); expect(body.clip.width).toBeCloseTo(80, 1);
  expect(body.body.x).toBeCloseTo(rtl ? 0 : -40, 1); expect(body.body.y).toBeCloseTo(-30, 1);
  expect(await page.evaluate(() => window.workbookGridHost.failedPaint())).toBe(true);
});

test('viewport DOM stays bounded at row 10,000 and the far worksheet corner', async ({ page }) => {
  const first = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: 0, scrollTopPt: 0 }); return window.workbookGridHost.inspect(); });
  const middle = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: 0, scrollTopPt: 99_970 }); return window.workbookGridHost.inspect(); });
  await expect(page.locator('[data-row="9999"][data-column="9"]')).toHaveText('一万行');
  const last = await page.evaluate(async () => { await window.workbookGridHost.render({ widthPt: 960, heightPt: 600, scrollLeftPt: Number.MAX_SAFE_INTEGER, scrollTopPt: Number.MAX_SAFE_INTEGER }); return window.workbookGridHost.inspect(); });
  await expect(page.locator('[data-row="1048575"][data-column="16383"]')).toHaveText('末格');
  for (const report of [first, middle, last]) {
    expect(report.cellCount).toBeLessThanOrEqual(5000); expect(report.sourceUnchanged).toBe(true);
    expect(report.borderLines).toBeGreaterThan(0); expect(report.borderLines).toBeLessThanOrEqual(report.cellCount * 4);
  }
  expect(last.cellCount).toBeLessThanOrEqual(48 * 60); expect(middle.cellCount).toBeLessThanOrEqual(48 * 60);
  await page.evaluate(() => window.workbookGridHost.dispose()); await expect(page.locator('#workbook-grid-probe')).toHaveCount(0);
});

test('General fitting and numeric fill agree with independently laid out fixed-font text', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.numericMeasurement());
  expect(result.rounded).toMatchObject({ kind: 'text', text: '26' });
  expect(result.scientific).toMatchObject({ kind: 'text', text: '1.23457E+11' });
  expect(result.overflow).toBe(false); expect(result.fill?.text).toMatch(/^f+$/);
  expect(result.actual).toBeLessThanOrEqual(140.25); expect(Math.abs(result.actual - result.predicted)).toBeLessThan(.25);
});

test('virtual grid renders formatted scalar values instead of imported display text in both directions', async ({ page }) => {
  for (const rtl of [false, true]) {
    const result = await page.evaluate(async rtl => {
      await window.workbookGridHost.render({ widthPt: 320, heightPt: 100, scrollLeftPt: 0, scrollTopPt: 4970 }, rtl);
      return window.workbookGridHost.inspect();
    }, rtl);
    await expect(page.locator('[data-office-view="workbook-cell"][data-row="500"][data-column="8"]')).toHaveText('12,345.68');
    expect(result.sourceUnchanged).toBe(true); expect(result.cellCount).toBeLessThanOrEqual(5000);
  }
});

test('numeric cell layouts format current values, fit real fonts and retain cell-based projection', async ({ page }) => {
  const cases: NumberTextCase[] = [
    { value: 1234.5, format: { code: '#,##0.00' }, widthPt: 90 },
    { value: -1.005, format: { code: '0.00;[Red](0.00)' }, widthPt: 80 },
    { value: 60, format: { builtinId: 14 }, widthPt: 100 },
    { value: 12.5, format: { code: '[$-407]0.00' }, widthPt: 80 },
    { value: .125, format: { code: '0.0%' }, widthPt: 80 },
    { value: 25.76, widthPt: 16 },
    { value: 12345.678, format: { code: '#,##0.00' }, widthPt: 20 },
    { value: -1, format: { builtinId: 14 }, widthPt: 80 },
    { value: 5, format: { code: '[>10]0;[<0]0' }, widthPt: 80 },
    { value: 12.5, format: { code: '"¥"* 0.00_)' }, widthPt: 100 },
    { value: 1.2345, widthPt: 15, styles: { '--office-cell-shrink-to-fit': 'true' } },
    { value: 12.5, widthPt: 75, heightPt: 75, styles: { '--office-cell-text-rotation': '45' } },
    { value: 12345, widthPt: 30, heightPt: 120, styles: { '--office-cell-text-rotation': '255' } },
    { value: 1, widthPt: 75, format: { code: '[Green]0.00' } },
    { value: 1, widthPt: 75, format: { code: '[Color3]0.00' }, styles: { '--office-indexed-10': '#123456' } },
    { value: 12.5, widthPt: 160, format: { code: '"预算 "0.0' } },
    { value: 1, widthPt: 75, format: { code: ';;' } },
  ];
  const report = await page.evaluate(cases => window.workbookGridHost.numberTextMeasurement(cases), cases), items = report.measurements;
  expect(items.slice(0, 6).map(item => item.text)).toEqual(['1,234.50', '(1.01)', '1900/2/29', '12,50', '12.5%', '26']);
  expect(items[1]!.color).toBe('rgb(255, 0, 0)'); expect(items[3]!.locale).toBe('de-DE');
  for (const index of [6, 7, 8]) expect(items[index]!.text).toMatch(/^#+$/);
  expect(items.slice(6, 9).map(item => item.overflow)).toEqual(['width', 'invalidDate', 'unmatchedCondition']);
  expect(items[9]!.text).toMatch(/^¥ +12\.50$/); expect(items[9]!.spaces).toHaveLength(1);
  expect(Math.abs(items[9]!.spaces[0]!.actualWidth - items[9]!.spaces[0]!.sourceWidth)).toBeLessThan(.1);
  expect(items[10]!.text).toBe('1.2345'); expect(items[10]!.scale).toBeLessThan(1);
  expect(items[11]!.transform).toContain('rotate(-45deg)'); expect(items[12]!.text).toBe('12345'); expect(items[12]!.breaks).toBe(4);
  expect(items[13]!.color).toBe('rgb(0, 255, 0)'); expect(items[14]!.color).toBe('rgb(18, 52, 86)');
  expect(items[15]!.text).toBe('预算 12.5'); expect(items[15]!.fonts).toContainEqual(expect.objectContaining({ role: 'east-asia', family: expect.stringMatching(/^"?Noto Sans CJK SC"?$/) }));
  expect(items[16]!.text).toBe('');
  for (const [index, item] of items.entries()) {
    expect(item.projections).toBe(0);
    if (![11, 12].includes(index)) { expect(item.widthPt).toBeLessThanOrEqual(cases[index]!.widthPt - 2.25 + .1); expect(Math.abs(item.rightGapPt - 1.125)).toBeLessThan(.1); }
  }
  expect(report.unchanged).toBe(true); expect(report.atomic).toBe(true); expect(report.retained).toBe(0);
  expect(report.failures).toEqual(['ResourceLimitExceeded', 'ResourceLimitExceeded', 'UnsupportedFeature', 'InvalidOperation', 'InvalidOperation']);
});

test('system date and time patterns share explicit locale semantics across runtime versions', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.numberTextMeasurement([
    { value: 45351.5, format: { code: '[$-F800]yyyy-mm-dd' }, widthPt: 240 },
    { value: .5, format: { code: '[$-F400]h:mm' }, widthPt: 100 },
    { value: .5, format: { code: 'h:mm"\u202f"AM/PM' }, widthPt: 100 },
  ], 'en-US'));
  expect(report.measurements.map(item => item.text)).toEqual(['Thursday, February 29, 2024', '12:00:00 PM', '12:00\u202fPM']);
  expect(report.unchanged).toBe(true); expect(report.retained).toBe(0);
});

test('column font metrics use their declared Normal font and keep legacy defaults isolated', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.columnFontMeasurement());
  expect(report.measured.cssClass).toBe('normal-metric'); expect(report.fallback.cssClass).toBe('cell-default');
  expect(report.measured.maximumDigitWidthPx).toBeGreaterThan(report.fallback.maximumDigitWidthPx * 1.4);
  expect(report.measured.fonts).toContainEqual(expect.objectContaining({ requestedFamily: 'Liberation Serif', substituted: false }));
  expect(Math.abs(report.actual - report.measured.maximumDigitWidthPx)).toBeLessThan(.1); expect(report.retainedProbes).toBe(0);
  expect(report.unrelatedUnchanged).toBe(true);
});

test('worksheet colors resolve theme variables and luminance without treating native ARGB bytes as opacity', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.colorMeasurement());
  expect(result.values).toEqual(['rgba(143,170,220,1)', 'rgba(51,90,161,1)', 'rgba(20,40,60,1)', 'rgba(255,0,0,1)', 'rgba(20,40,60,0.5)', 'rgba(0,255,0,1)', undefined]);
  expect(result.errors).toEqual(['InvalidDocument', 'InvalidDocument', 'InvalidDocument']); expect(result.unchanged).toBe(true);
  expect(result.fills).toEqual(['rgb(1, 2, 3)', 'transparent', 'rgba(143,170,220,1)', 'pattern']); expect(result.fillErrors).toEqual(['InvalidDocument']);
});

test('native cell fills match standard patterns and preserve full-cell gradients across clipped panes', async ({ page }) => {
  const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync('packages/office/tests/fixtures/xlsx-cell-appearance.xlsx')));
  const inputs = appearanceReference.cases.filter(input => input.kind === 'pattern' || input.kind === 'gradient');
  const classes = inputs.map((_, index) => model.getCell(`sheet-${index + 1}`, 1, 1)!.cssClass!);
  const report = await page.evaluate(({ css, classes }) => window.workbookGridHost.measureWorkbookFills(css, classes), { css: model.snapshot().styles.map(style => style.css).join('\n'), classes });
  expect(report.reports).toHaveLength(27);
  expect(report.invalid).toEqual(['InvalidDocument', 'ResourceLimitExceeded', 'InvalidDocument', 'InvalidDocument']);
  for (const [index, actual] of report.reports.entries()) {
    const input = inputs[index]!;
    expect(actual.sourceUnchanged).toBe(true); expect(actual.immutable).toBe(true); expect(actual.nodes).toBe(actual.actual);
    expect(actual.extent).toEqual(['24pt', '18pt']); expect(actual.largeNodes).toBeLessThanOrEqual(actual.nodes);
    expect(actual.errors).toEqual(['ResourceLimitExceeded']); expect(actual.difference, JSON.stringify(input)).toBeLessThanOrEqual(3);
    if (input.kind === 'pattern') {
      if (input.pattern === 'none') { expect(actual.samples.every(sample => sample.color[3] === 0)).toBe(true); continue; }
      if (input.pattern === 'solid') { expect(actual.samples.every(sample => sample.color.join(',') === '32,64,128,255')).toBe(true); continue; }
      const expected = fillReference.cases.find(item => item.pattern === input.pattern)!;
      const rows = expected.rows, width = rows[0]!.length, height = rows.length;
      let matches = false;
      for (let dy = 0; dy < height; dy++) for (let dx = 0; dx < width; dx++) if (actual.rows.every((row, y) => [...row].every((pixel, x) => pixel === rows[(y + dy) % height]![(x + dx) % width]))) matches = true;
      expect(matches, input.pattern).toBe(true);
      expect(actual.rows.join('').split('1').length - 1).toBe(expected.foregroundFraction * 1024);
    } else {
      for (const sample of actual.samples) {
        const x = (sample.x + .5) / 96, y = (sample.y + .5) / 72;
        let fraction: number;
        if (input.type === 'linear') {
          const angle = input.degree! * Math.PI / 180, dx = Math.cos(angle), dy = Math.sin(angle);
          fraction = .5 + ((x - .5) * 96 * dx + (y - .5) * 72 * dy) / (96 * Math.abs(dx) + 72 * Math.abs(dy));
        } else {
          const left = input.left!, right = input.right!, top = input.top!, bottom = input.bottom!;
          fraction = Math.max(0, x < left ? (left - x) / left : x > right ? (x - right) / (1 - right) : 0, y < top ? (top - y) / top : y > bottom ? (y - bottom) / (1 - bottom) : 0);
        }
        const expected = fraction <= .5 ? [255 * (1 - 2 * fraction), 510 * fraction, 0, 255] : [0, 510 * (1 - fraction), 255 * (2 * fraction - 1), 255];
        for (let channel = 0; channel < 4; channel++) expect(Math.abs(sample.color[channel]! - Math.round(expected[channel]!)), `${JSON.stringify(input)} at ${sample.x},${sample.y} channel ${channel}`).toBeLessThanOrEqual(3);
      }
    }
  }
});

test('native diagonal borders preserve direction, double spacing and standard slant-dash-dot strokes', async ({ page }) => {
  const { model } = projectXlsxRecords(await readXlsxRecords(readFileSync('packages/office/tests/fixtures/xlsx-cell-appearance.xlsx')));
  const inputs = appearanceReference.cases.map((input, index) => ({ input, cssClass: model.getCell(`sheet-${index + 1}`, 1, 1)!.cssClass! })).filter(({ input }) => input.kind === 'diagonal');
  const report = await page.evaluate(({ css, classes }) => window.workbookGridHost.diagonalMeasurement(css, classes), { css: model.snapshot().styles.map(style => style.css).join('\n'), classes: inputs.map(input => input.cssClass) });
  expect(report.result).toHaveLength(9); expect(report.failures).toEqual(Array(4).fill('InvalidDocument'));
  for (const [index, actual] of report.result.entries()) {
    const input = inputs[index]!.input, perDirection = input.style === 'thin' ? 1 : 2;
    expect(actual.unchanged).toBe(true); expect(actual.difference, JSON.stringify(input)).toBeLessThanOrEqual(2); expect(actual.maskDifference, JSON.stringify(input)).toBeLessThanOrEqual(1);
    expect(actual.strokes).toHaveLength((Number(input.up) + Number(input.down)) * perDirection);
    expect(actual.strokes.filter(line => line.to[1]! < line.from[1]!)).toHaveLength(input.up ? perDirection : 0);
    expect(actual.strokes.filter(line => line.to[1]! > line.from[1]!)).toHaveLength(input.down ? perDirection : 0);
    for (const line of actual.strokes) { expect(line.color).toBe('rgba(0,68,170,1)'); expect(line.width).toBe(input.style === 'double' ? .5 : .75); }
    if (input.style === 'double') {
      const [a, b] = actual.strokes; const dx = a!.to[0]! - a!.from[0]!, dy = a!.to[1]! - a!.from[1]!;
      expect(Math.abs((b!.from[0]! - a!.from[0]!) * dy - (b!.from[1]! - a!.from[1]!) * dx) / Math.hypot(dx, dy)).toBeCloseTo(1.25);
    }
  }
  const width = borderSpec.rows[0]!.length;
  expect(Array.from({ length: width }, (_, phase) => report.rows.every((row, y) => [...row].every((pixel, x) => pixel === borderSpec.rows[y]![(x + phase) % width]))).some(Boolean)).toBe(true);
});

test('bounded cell style projection retains source chunk selectors, rich runs and lifecycle isolation', async ({ page }) => {
  const requests: string[] = []; page.on('request', request => { if (request.url().includes('workbook-injected')) requests.push(request.url()); });
  const report = await page.evaluate(() => window.workbookGridHost.sourceStyleMeasurement());
  expect(report).toMatchObject({ text: '一万行', italic: 'italic', background: 'rgb(68, 114, 196)', bold: '700', foreground: 'rgb(153, 51, 85)', plain: 'cell-default', blockRow: '9984', row: '9999', bound: true, disposed: true, noPartial: true, retained: 0 });
  expect(report.font).toContain('Liberation Serif'); expect(parseFloat(report.size)).toBeCloseTo(13 / .75, 1);
  expect(report.before).toContain('<span class="rich-style">一万</span>行');
  expect(report.rejected).toEqual(Array(6).fill('InvalidDocument')); expect(requests).toEqual([]);
  expect(await page.evaluate(() => (window as unknown as Record<string, unknown>).workbookInjected)).toBeUndefined();
});

test('cell font roles retain UTF-16 and variation selectors across source spans without changing CSS ancestry', async ({ page }) => {
  const result = await page.evaluate(() => window.workbookGridHost.textFontMeasurement());
  expect(result).toMatchObject({ limited: true, unchanged: true, frozen: true });
  expect(result.runs.map(({ text, from, to, role }) => ({ text, from, to, role }))).toEqual([
    { text: '合并', from: 0, to: 2, role: 'east-asia' }, { text: '😀', from: 2, to: 4, role: 'h-ansi' },
    { text: '禰', from: 4, to: 5, role: 'east-asia' }, { text: '󠄀', from: 5, to: 7, role: 'east-asia' },
  ]);
  for (const run of result.runs.filter(run => run.role === 'east-asia')) expect(run.font.family).toContain('Noto Sans CJK SC');
  expect(result.rich.map(run => run.font.weight)).toEqual(['700', '400']);
  for (const run of result.rich) { expect(run.font.style).toBe('italic'); expect(run.font.family).toContain('Liberation Serif'); expect(run.font.sizePx).toBeCloseTo(13 / .75, 1); }
  expect(result.withBreak.map(run => run.text).join('')).toBe('合并😀禰󠄀\nA');
  expect(result.withBreak.at(-1)).toMatchObject({ text: 'A', from: 8, to: 9, role: 'ascii' });
});

test('rich text display preserves source offsets, local style and independent native decorations', async ({ page }, testInfo) => {
  const result = await page.evaluate(() => window.workbookGridHost.textViewMeasurement());
  expect(result).toMatchObject({ text: 'A\nB', unchanged: true, frozen: true,
    decoration: { strike: 'line-through', strikeStyle: 'solid', underline: 'underline', underlineStyle: 'double', plain: 'none' },
    failures: ['ResourceLimitExceeded', 'UnsupportedFeature', 'InvalidDocument'] });
  expect(result.rich.map(run => [run.text, run.from, run.to, run.weight, run.color])).toEqual([['一万', 0, 2, '700', 'rgb(153, 51, 85)'], ['行', 2, 3, '400', 'rgb(0, 0, 0)']]);
  expect(result.rich[0]!.x + result.rich[0]!.width).toBeCloseTo(result.rich[1]!.x, 0);
  expect(result.mixed.map(run => [run.text, run.from, run.to])).toEqual([['合并', 0, 2], ['😀', 2, 4], ['禰', 4, 5], ['󠄀', 5, 7]]);
  for (const run of result.mixed.filter(run => run.text !== '😀')) expect(run.family).toContain('Noto Sans CJK SC');
  expect(result.decorated.map(run => [run.text, run.from, run.to, run.color])).toEqual([['A', 0, 1, 'rgb(143, 170, 220)'], ['\n', 1, 2, 'rgb(143, 170, 220)'], ['B', 2, 3, 'rgb(143, 170, 220)']]);
  expect(result.decorated[2]!.y).toBeGreaterThan(result.decorated[0]!.y + result.decorated[0]!.height * .8);
  const screenshot = testInfo.outputPath('workbook-rich-text.png'); await page.locator('#workbook-text-view-probe').screenshot({ path: screenshot });
  await testInfo.attach('workbook-rich-text', { path: screenshot, contentType: 'image/png' });
});

test('border strokes match native point geometry and keep dash phase across clipped panes', async ({ page }, testInfo) => {
  const native = Object.entries(borderReference.kinds.pdfStrokes).filter(([kind]) => kind !== 'none' && kind !== 'slantDashDot');
  const edges = native.map(([kind, strokes]) => ({ kind: kind as WorkbookBorderKind, color: 'rgba(255,0,0,1)',
    from: { x: strokes[0]!.from[0]!, y: strokes.reduce((sum, stroke) => sum + stroke.from[1]!, 0) / strokes.length },
    to: { x: strokes[0]!.to[0]!, y: strokes.reduce((sum, stroke) => sum + stroke.to[1]!, 0) / strokes.length } }));
  const report = await page.evaluate(edges => window.workbookGridHost.borderViewMeasurement(edges), edges);
  expect(report.strokes).toHaveLength(native.flatMap(([, strokes]) => strokes).length);
  let index = 0;
  for (const [, strokes] of native) for (const expected of [...strokes].sort((a, b) => a.from[1]! - b.from[1]!)) {
    const actual = report.strokes[index++]!;
    expect(Math.abs(actual.width - expected.widthPt)).toBeLessThan(.02); expect(actual.cap).toBe(expected.cap);
    expect(actual.dash).toHaveLength(expected.dashPt.length);
    for (let n = 0; n < actual.dash.length; n++) expect(Math.abs(actual.dash[n]! - expected.dashPt[n]!)).toBeLessThan(.02);
    for (const point of ['from', 'to'] as const) for (let axis = 0; axis < 2; axis++) expect(Math.abs(actual[point][axis]! - expected[point][axis]!)).toBeLessThan(.01);
  }
  for (let i = 0; i < report.rasterImages.length; i++) await testInfo.attach(`border-raster-${i}`, { body: Buffer.from(report.rasterImages[i]!.split(',')[1]!, 'base64'), contentType: 'image/png' });
  expect(report.rasterSizes).toEqual(Array(4).fill([160, 80]));
  expect(report.splitMaskDifference).toBe(0); expect(report.splitDifference).toBeLessThanOrEqual(1);
  expect(report.maskDifference).toBe(0); expect(report.difference).toBeLessThanOrEqual(1); expect(report.painted).toBeGreaterThan(70); expect(report.painted).toBeLessThan(110);
  expect(report.doubleRows.length).toBeGreaterThanOrEqual(2);
  expect(report.doubleRows.some((row, index) => index && row - report.doubleRows[index - 1]! >= 3)).toBe(true);
  expect(report.failures).toEqual(['UnsupportedFeature', 'InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded']);
  const screenshot = testInfo.outputPath('workbook-border-strokes.png'); await page.locator('#workbook-border-view-probe').screenshot({ path: screenshot });
  await testInfo.attach('native-calibrated-border-strokes', { path: screenshot, contentType: 'image/png' });
});

test('worksheet context reading order uses its first non-whitespace scalar with isolated Unicode direction', async ({ page }) => {
  const report = await page.evaluate(() => window.workbookGridHost.textDirectionMeasurement());
  expect(report).toEqual({ values: ['ltr', 'rtl', 'ltr', 'ltr', 'rtl', 'ltr', 'ltr', 'rtl', 'rtl', 'ltr', 'rtl'], explicit: ['ltr', 'rtl'], wholeStringDirection: 'rtl',
    retainedText: '', failures: ['InvalidDocument', 'InvalidOperation', 'InvalidOperation'], remaining: 0, ownerDirection: 'rtl' });
});

test('worksheet text layout measures native alignment, Normal indentation and local fonts before clipping', async ({ page }, testInfo) => {
  const cases: TextLayoutCase[] = alignmentReference.cases.map(item => {
    const font = alignmentReference.fonts[item.font]!;
    return { text: item.kind === 'b' ? 'TRUE' : String(item.value), type: item.kind === 'n' ? 'number' : item.kind === 'b' ? 'boolean' : item.kind === 'e' ? 'error' : 'text',
      family: font.family, sizePt: font.sizePt, widthPt: item.cellBoundsPt[2]! - item.cellBoundsPt[0]!, heightPt: item.cellBoundsPt[3]! - item.cellBoundsPt[1]!,
      horizontal: item.horizontal, vertical: item.vertical, indent: item.indent, readingOrder: ['context', 'ltr', 'rtl'][item.readingOrder]!, rightToLeft: item.rightToLeft };
  });
  const result = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  await testInfo.attach('alignment-measurements', { body: JSON.stringify(result, null, 2), contentType: 'application/json' });
  expect(result).toMatchObject({ unchanged: true, retained: 0, atomic: true, frozen: true,
    failures: ['UnsupportedFeature', 'UnsupportedFeature', 'InvalidDocument', 'InvalidDocument', 'InvalidDocument', 'InvalidOperation', 'InvalidOperation', 'ResourceLimitExceeded', 'InvalidDocument'] });
  for (const [index, item] of cases.entries()) {
    const measured = result.measurements[index]!, native = alignmentReference.cases[index]!;
    expect(measured.text).toBe(item.text); expect(measured.scale).toBe(1);
    const direction = item.type === 'text' && (item.readingOrder === 'rtl' || item.readingOrder === 'context' && item.text === 'אבג') ? 'rtl' : 'ltr';
    expect(measured.direction).toBe(direction);
    if (item.type === 'error' || item.text.startsWith('-')) expect(measured.firstX).toBeLessThan(measured.lastX);
    const horizontal = item.horizontal !== 'general' ? item.horizontal : item.type === 'number' ? 'right' : item.type === 'boolean' || item.type === 'error' ? 'center' : direction === 'rtl' ? 'right' : 'left';
    expect(measured.horizontal).toBe(horizontal);
    const verticalSide = item.vertical === 'bottom' ? 3 : 1;
    expect(Math.abs(measured.gapsPt[verticalSide]! - native.textGapsPt[verticalSide]!), `vertical case ${index}`).toBeLessThan(.8);
    // Explicit placements and Latin General can be compared with this producer.
    // Its Boolean/error alignment and explicit-LTR Hebrew differ from OOXML.
    if (item.type === 'boolean' || item.type === 'error' || item.horizontal === 'general' && item.text === 'אבג' && item.readingOrder === 'ltr') continue;
    const aligned = horizontal === 'right' ? 2 : 0;
    // Calc's RTL PDF places the same content about .44pt farther right relative
    // to its red cell border. Record that producer difference separately; the
    // SDK must keep explicit physical alignment invariant under sheet reversal.
    if (item.rightToLeft) {
      const ltr = result.measurements[index - cases.length / 2]!;
      for (let side = 0; side < 4; side++) expect(measured.gapsPt[side]).toBeCloseTo(ltr.gapsPt[side]!, 1);
    } else if (horizontal !== 'center') expect(Math.abs(measured.gapsPt[aligned]! - native.textGapsPt[aligned]!), `horizontal case ${index}`).toBeLessThan(.65);
    else expect(Math.abs(measured.gapsPt[0]! - measured.gapsPt[2]!)).toBeLessThan(.1);
    expect(Math.abs(measured.widthPt - (native.cellBoundsPt[2]! - native.cellBoundsPt[0]! - native.textGapsPt[0]! - native.textGapsPt[2]!)), `text width case ${index}`).toBeLessThan(.15);
  }
  expect(result.fonts.every(font => !font.substituted)).toBe(true);
  await page.locator('#workbook-text-layout-probe').screenshot({ path: testInfo.outputPath('workbook-text-alignment.png') });
});

test('rotated and stacked text matches native ink bounds while retaining complete graphemes', async ({ page }, testInfo) => {
  const cases: TextLayoutCase[] = rotationReference.cases.map(item => ({ text: item.text, type: 'text', family: rotationReference.font.family, sizePt: rotationReference.font.sizePt,
    widthPt: item.cellSizePt[0]!, heightPt: item.cellSizePt[1]!, horizontal: 'center', vertical: 'center', indent: 0, readingOrder: 'context', rightToLeft: false, rotation: item.rotation }));
  cases.push({ ...cases[12]!, splitOffsets: [1, 3, 4, 5] });
  const report = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), cases);
  expect(report).toMatchObject({ unchanged: true, retained: 0, atomic: true, frozen: true, usedFallbackLineMetrics: false,
    stackFailures: ['InvalidDocument', 'InvalidDocument', 'ResourceLimitExceeded'] });
  const inks = [];
  for (const [index, input] of cases.entries()) {
    const measured = report.measurements[index]!;
    expect(measured.text).toBe(input.text); expect(measured.scale).toBe(1); expect(measured.offsets[0]![0]).toBe(0); expect(measured.offsets.at(-1)![1]).toBe(input.text.length);
    for (let i = 1; i < measured.offsets.length; i++) expect(measured.offsets[i]![0]).toBe(measured.offsets[i - 1]![1]);
    if (input.rotation === 255) expect(measured.breaks).toBe(index >= 12 ? 2 : index === 10 ? 2 : 4);
    else if (input.rotation) expect(measured.transform).toBe(`rotate(${input.rotation <= 90 ? -input.rotation : input.rotation - 90}deg) scale(1)`);
    const bytes = await page.locator('#workbook-text-layout-probe > div').nth(index).screenshot();
    const ink = await page.evaluate(async base64 => {
      const image = new Image(); image.src = 'data:image/png;base64,' + base64; await image.decode();
      const canvas = document.createElement('canvas'); canvas.width = image.width; canvas.height = image.height; const context = canvas.getContext('2d')!; context.drawImage(image, 0, 0);
      const pixels = context.getImageData(0, 0, image.width, image.height).data;
      return [64, 224].map(threshold => {
        let left = Infinity, top = Infinity, right = 0, bottom = 0;
        for (let y = 0; y < image.height; y++) for (let x = 0; x < image.width; x++) {
          const offset = (y * image.width + x) * 4;
          // Include LCD subpixel antialiasing. The gallery's pure-red outline
          // already lies outside the intensity threshold; filtering colored
          // edge pixels would remove real glyph coverage in Chromium.
          const maximum = Math.max(pixels[offset]!, pixels[offset + 1]!, pixels[offset + 2]!);
          if (pixels[offset + 3]! >= 128 && maximum < threshold) { left = Math.min(left, x); top = Math.min(top, y); right = Math.max(right, x + 1); bottom = Math.max(bottom, y + 1); }
        }
        return [left, top, right, bottom].map(value => value * .75);
      });
    }, bytes.toString('base64'));
    inks.push(ink);
    // The producer splits UTF-16 units in the last case, including surrogates
    // and combining marks. Its corrupted glyph boxes are not a reference.
    if (index < 12) for (let side = 0; side < 4; side++) expect.soft(Math.abs(ink[1]![side]! - rotationReference.cases[index]!.coverageBoundsPt[side]!), `rotation ${input.rotation}, case ${index}, side ${side}: ${JSON.stringify(ink)}`).toBeLessThanOrEqual(2);
  }
  await testInfo.attach('rotation-measurements', { body: JSON.stringify({ report, inks }, null, 2), contentType: 'application/json' });
  await testInfo.attach('rotation-gallery', { body: await page.locator('#workbook-text-layout-probe').screenshot(), contentType: 'image/png' });
});

test('worksheet wrapping and shrink preserve text offsets and report missing fonts without changing source', async ({ page }) => {
  const base: TextLayoutCase = { text: 'Long text 汉字𠮷禰󠄀', type: 'text', family: 'Liberation Sans', sizePt: 11, widthPt: 36, heightPt: 100,
    horizontal: 'left', vertical: 'top', indent: 0, readingOrder: 'context', rightToLeft: false };
  const result = await page.evaluate(cases => window.workbookGridHost.textLayoutMeasurement(cases), [
    base, { ...base, shrink: true }, { ...base, wrap: true, shrink: true }, { ...base, text: 'Long line\nSecond line', shrink: true },
    { ...base, family: 'Missing Workbook Layout Face', widthPt: 200 }, { ...base, text: 'A', vertical: 'bottom', widthPt: 100 }, { ...base, text: 'Long line', widthPt: 100 },
  ]);
  expect(result).toMatchObject({ unchanged: true, retained: 0, atomic: true });
  const [plain, shrunk, wrapped, multiline] = result.measurements;
  expect(shrunk!.scale).toBeLessThan(1); expect(shrunk!.widthPt).toBeLessThanOrEqual(36 - 2.25 + .1);
  expect(wrapped!.scale).toBe(1); expect(wrapped!.requiredHeightPt).toBeGreaterThan(plain!.requiredHeightPt * 2);
  expect(wrapped!.widthPt).toBeLessThanOrEqual(36 - 2.25 + .1); expect(multiline!.scale).toBe(1);
  expect(multiline!.requiredHeightPt).toBeCloseTo((result.measurements[6]!.requiredHeightPt - 1.5) * 2 + 1.5, 1);
  for (const item of [plain!, shrunk!, wrapped!]) { expect(item.text).toBe(base.text); expect(item.offsets.at(-1)![1]).toBe(base.text.length); }
  expect(result.fonts.find(font => font.requestedFamily === 'Missing Workbook Layout Face')).toMatchObject({ substituted: true });
});

test('source border resolution keeps native logical sides and physical CSS stroke widths distinct', async ({ page }, testInfo) => {
  const result = await page.evaluate(() => window.workbookGridHost.borderSourceMeasurement());
  expect(result).toMatchObject({ unchanged: true, failures: ['InvalidDocument', 'UnsupportedFeature', 'UnsupportedFeature', 'UnsupportedFeature', 'InvalidDocument'] });
  expect(result.native[0]).toEqual(result.native[1]);
  expect(result.native[0]).toEqual({ before: { kind: 'thin', color: 'rgba(143,170,220,1)' }, after: { kind: 'double', color: 'rgba(10,20,30,1)' } });
  expect(result.css[0]!.before).toEqual({ kind: 'css', style: 'double', widthPt: 3, color: 'rgba(10,20,30,0.5)' });
  expect(result.css[0]!.after).toEqual({ kind: 'css', style: 'dashed', widthPt: 1.5, color: 'rgba(0,0,255,1)' });
  expect(result.css[0]!.top).toEqual({ kind: 'css', style: 'dotted', widthPt: .75, color: 'rgba(0,128,0,1)' });
  expect(result.css[1]!.after).toEqual(result.css[0]!.before); expect(result.css[1]!.before).toEqual(result.css[0]!.after);
  expect(result.aliases[0]!.before!.kind).toBe('medium'); expect(result.aliases[0]!.after!.kind).toBe('thin');
  expect(result.aliases[1]).toEqual(result.aliases[2]);
  expect(result.strokes.map(item => [item.width, item.dash, item.cap])).toEqual([
    ['1', null, 'butt'], ['1', null, 'butt'], ['1.5', '4.5 1.5', 'butt'], ['0.75', '0 1.5', 'round'], ['0.5', null, 'butt'], ['0.5', null, 'butt'],
  ]);
  await page.locator('#workbook-css-border-probe').screenshot({ path: testInfo.outputPath('workbook-css-borders.png') });
});

test('workbook asset leases decode authorized images and isolate URL lifetimes without copying cell chunks', async ({ page }) => {
  for (const { bytes, mimeType, frames } of workbookRasterCases()) {
    const result = await page.evaluate(({ base64, mimeType }) => window.workbookGridHost.assetMeasurement(base64, mimeType), { base64: bytes.toString('base64'), mimeType });
    expect(result).toEqual({ count: 1, byteLength: bytes.length, pixels: 64 * frames, distinct: true,
      dimensions: Array(2).fill([8, 8]), revoked: true, isolated: true, failure: 'OperationCancelled' });
  }
  await expect(page.locator('#workbook-grid-probe img')).toHaveCount(0);
});
