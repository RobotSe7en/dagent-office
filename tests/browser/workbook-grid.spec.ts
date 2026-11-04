import { expect, test } from "@playwright/test";
import { build } from "vite";
import { fileURLToPath } from "node:url";
import type { WorkbookViewport } from "../../packages/render/src/workbook/viewport.js";

interface GridHost {
  render(viewport: WorkbookViewport, rtl?: boolean): Promise<boolean>;
  inspect(): { cellCount: number; accessibleCells: number; rows: number[]; sourceUnchanged: boolean; merges: { pane?: string; text: string | null; clip: { x: number; y: number; width: number; height: number }; body: { x: number; y: number; width: number; height: number }; hit?: { row: number; column: number } }[] };
  failedPaint(): boolean; dispose(): void;
}
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
  for (const report of [first, middle, last]) { expect(report.cellCount).toBeLessThanOrEqual(5000); expect(report.sourceUnchanged).toBe(true); }
  expect(last.cellCount).toBeLessThanOrEqual(48 * 60); expect(middle.cellCount).toBeLessThanOrEqual(48 * 60);
  await page.evaluate(() => window.workbookGridHost.dispose()); await expect(page.locator('#workbook-grid-probe')).toHaveCount(0);
});
