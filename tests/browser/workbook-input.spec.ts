import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { openSdkHost } from "./sdk-host.js";
import type * as fixture from "./fixtures/workbook-input.js";
import type * as editorFixture from "./fixtures/workbook-editor.js";
import { scalarWorkbook } from "../../packages/core/tests/workbook-scalar-fixture.js";
import { encodeTestPackage } from "./package-transport.js";
declare global { interface Window { workbookInput: typeof fixture } }
declare global { interface Window { workbookEditor: typeof editorFixture } }
let code = "", editorCode = "";
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/workbook-input.ts", import.meta.url)), formats: ["iife"], name: "WorkbookInput" } } });
  const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
  code = output.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n");
  const editor = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/workbook-editor.ts", import.meta.url)), formats: ["iife"], name: "WorkbookEditorFixture" } } });
  const chunks = Array.isArray(editor) ? editor.flatMap(part => part.output) : "output" in editor ? editor.output : [];
  editorCode = chunks.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n");
});
function editorSource(rtl = false, structure = true) {
  return encodeTestPackage(scalarWorkbook([{ row: 0, column: 0, value: { type: "number", value: 1 } }, { row: 0, column: 1, value: { type: "number", value: 999 }, formula: "=A1+1" }], metadata => {
    const sheet = metadata.workbook.sheets[0]!; sheet.rightToLeft = rtl; sheet.defaultRowHeightCustom = true;
    if (structure) {
      sheet.rowDimensions = [{ from: 1, to: 1000, hidden: true }]; sheet.columnDimensions = [{ from: 1, to: 10, unit: "characters", hidden: true }];
      sheet.merges = [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }];
    }
    metadata.workbook.sheets.push({ ...sheet, sheetId: "second", name: "Second", rowDimensions: [], columnDimensions: [], merges: [] });
  }).serialize());
}
test("mounted workbook formula input, calculation, address selection and toolbar history share source", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource());
  await page.evaluate(() => window.workbookEditor.select(0)); const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true });
  await input.pressSequentially("5"); expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 } });
  await input.press("Enter"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, cleanSource: true, errors: [],
    calculated: { state: "valid", value: { type: "number", value: 6 } }, selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }] } } });
  const address = page.getByRole("textbox", { name: "选择单元格或区域", exact: true }); await address.fill("A1"); await address.press("Enter"); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("5");
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, calculated: { value: { type: "number", value: 2 } }, errors: [] });
  await page.getByRole("button", { name: "重做", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("5");
});
for (const rtl of [false, true]) test(`mounted workbook keyboard extends source ranges across hidden dimensions in ${rtl ? "RTL" : "LTR"}`, async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(rtl)); await page.evaluate(() => window.workbookEditor.select(0));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, selection: { anchor: { ranges: [{ from: { row: 0, column: 11 }, to: { row: 0, column: 11 } }] } } });
  await input.press(rtl ? "ArrowRight" : "ArrowLeft"); await page.evaluate(() => window.workbookEditor.settled()); await input.press("ArrowDown"); await page.evaluate(() => window.workbookEditor.settled());
  await input.press("Shift+ArrowDown"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, errors: [], selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1003, column: 2 } }] } } });
  await input.press("Shift+ArrowUp"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ selection: { anchor: { ranges: [{ from: { row: 1001, column: 0 }, to: { row: 1002, column: 2 } }] } } });
  await page.evaluate(() => window.workbookEditor.select(0));
  await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await input.press(rtl ? "ArrowLeft" : "ArrowRight"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, selection: { anchor: { ranges: [{ from: { row: 0, column: 12 }, to: { row: 0, column: 12 } }] } } });
});
test("mounted workbook mouse and sheet switching retain invalid drafts until explicit cancellation", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(false, false)); await page.evaluate(() => window.workbookEditor.select(1));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }); await input.fill("=");
  const point = await page.evaluate(() => window.workbookEditor.point(0, "A")); await page.mouse.click(point.x, point.y); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, input: { text: "=", dirty: true }, selection: { anchor: { ranges: [{ from: { row: 1, column: 0 }, to: { row: 1, column: 0 } }] } }, errors: ["InvalidOperation"] });
  await page.getByRole("tab", { name: "Second", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled()); await expect(input).toHaveValue("=");
  await input.press("Escape"); await expect(input).toHaveValue(""); await page.getByRole("tab", { name: "Second", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ input: { text: "" }, selection: { anchor: { sheetId: "second" } }, revision: 0 });
  await input.pressSequentially("new sheet"); await input.press("Control+Enter"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("second", 0, 0))).toMatchObject({ value: { type: "text", value: "new sheet" } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 0, 0))).toMatchObject({ value: { type: "number", value: 1 } });
});
test("mounted workbook TSV paste and rectangle clear use single cross-chunk history steps", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource(false, false)); await page.evaluate(() => window.workbookEditor.select(255, 31));
  expect(await page.evaluate(() => window.workbookEditor.paste("00123\t=AF256+1\n中文😀禰󠄀\tTRUE"))).toBe(true); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 1 }, cleanSource: true, errors: [], selection: { anchor: { ranges: [{ from: { row: 255, column: 31 }, to: { row: 256, column: 32 } }] } } });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toMatchObject({ value: { type: "text", value: "中文😀禰󠄀" } });
  await page.getByRole("textbox", { name: "单元格或公式输入", exact: true }).press("Delete"); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ history: { undoSteps: 2 }, errors: [] });
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toBeUndefined();
  await page.getByRole("button", { name: "撤销", exact: true }).click(); await page.evaluate(() => window.workbookEditor.settled());
  expect(await page.evaluate(() => window.workbookEditor.cell("sheet-1", 256, 31))).toMatchObject({ value: { type: "text", value: "中文😀禰󠄀" } });
});
test("mounted workbook real IME preserves a conflicting draft and disposal leaves its host session usable", async ({ page, browserName }) => {
  test.skip(browserName !== "chromium", "CDP composition is Chromium-specific.");
  await page.evaluate(() => window.workbookInput.dispose()); await page.addScriptTag({ content: editorCode }); await page.evaluate(entries => window.workbookEditor.mount(entries), editorSource()); await page.evaluate(() => window.workbookEditor.select(0));
  const input = page.getByRole("textbox", { name: "单元格或公式输入", exact: true }), cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 }); expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 0, input: { composing: true } });
  await page.evaluate(() => window.workbookEditor.external("external")); await page.evaluate(() => window.workbookEditor.settled());
  await cdp.send("Input.insertText", { text: "中文😀" }); await input.press("Enter");
  expect(await page.evaluate(() => window.workbookEditor.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, input: { text: "中文😀", dirty: true, selection: { baseRevision: 0 } }, cells: [{ value: { type: "text", value: "external" } }, expect.anything()], errors: ["RevisionConflict"] });
  await input.press("Escape"); await expect(input).toHaveValue("external");
  expect(await page.evaluate(() => window.workbookEditor.dispose())).toEqual({ revision: 1, usable: true });
});
test.beforeEach(async ({ page }) => { await openSdkHost(page); await page.addScriptTag({ content: code }); await page.evaluate(() => window.workbookInput.mount()); });
test("native typing remains transient until one Enter transaction, with shared undo and redo", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true });
  await input.pressSequentially("00123"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { text: "00123", dirty: true } });
  await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 1, history: { undoSteps: 1 }, state: { dirty: false }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "00123" } }) })]), navigation: [{ direction: "down", extend: false }], errors: [] });
  await input.press("Control+z"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 0, redoSteps: 1 }, state: { text: "" }, errors: [] });
  await input.press("Control+y"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "00123" }, errors: [] });
  await input.press("F2"); await input.press("End"); await input.pressSequentially("X"); await input.press("Escape"); await expect(input).toHaveValue("00123");
  await input.press("Shift+ArrowRight"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, navigation: [{ direction: "down", extend: false }, { direction: "right", extend: true }] });
});
test("native formula and multiline editing preserve draft failures and source types", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true });
  await input.pressSequentially("="); await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, state: { text: "=", dirty: true }, errors: [{ code: "InvalidOperation", text: "=" }] });
  await input.pressSequentially("1+2"); await input.press("Control+Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, navigation: [], cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ formula: expect.objectContaining({ expression: "=1+2" }) }) })]) });
  await input.press("Control+a"); await input.pressSequentially("中文"); await input.press("Alt+Enter"); await input.pressSequentially("混排"); await input.press("Tab");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { text: "中文\n混排" }, navigation: [{ direction: "right", extend: false }], cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "中文\n混排" } }) })]) });
});
test("host-authorized TSV paste crosses chunks atomically, rejects stale input and clears a rectangle", async ({ page }) => {
  await page.evaluate(() => window.workbookInput.mount(255, 31, 256, 32));
  expect(await page.evaluate(() => window.workbookInput.paste("00123\t=AF256+1\n中文😀禰󠄀\tTRUE"))).toBe(true);
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, errors: [], cells: [expect.anything(),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "00123" } }) }),
    expect.objectContaining({ cell: expect.objectContaining({ formula: expect.objectContaining({ expression: "=AF256+1" }) }) }),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "中文😀禰󠄀" } }) }),
    expect.objectContaining({ cell: expect.objectContaining({ value: { type: "boolean", value: true } }) })] });
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await input.press("Delete");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, cells: Array.from({ length: 5 }, () => expect.not.objectContaining({ cell: expect.anything() })) });
  await input.press("Control+z"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "00123" } });
  await page.evaluate(() => window.workbookInput.external("external")); await page.evaluate(() => window.workbookInput.paste("stale"));
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { selection: { baseRevision: 3 } }, errors: [{ code: "RevisionConflict", text: "00123" }] });
});
test("real Chromium IME commits only on confirmation and retains text after a source conflict", async ({ page, browserName }) => {
  test.skip(browserName !== "chromium", "CDP composition is Chromium-specific; all engines cover native text input and history.");
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }), cdp = await page.context().newCDPSession(page);
  await cdp.send("Input.imeSetComposition", { text: "中", selectionStart: 1, selectionEnd: 1 });
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { composing: true } });
  await cdp.send("Input.imeSetComposition", { text: "中文", selectionStart: 2, selectionEnd: 2 }); await cdp.send("Input.insertText", { text: "中文😀禰󠄀" });
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, state: { composing: false, text: "中文😀禰󠄀" } });
  await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "中文😀禰󠄀" } });
  await cdp.send("Input.imeSetComposition", { text: "迟到", selectionStart: 2, selectionEnd: 2 }); await page.evaluate(() => window.workbookInput.external("external")); await cdp.send("Input.insertText", { text: "迟到" }); await input.press("Enter");
  expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 2 }, state: { text: "迟到", dirty: true }, cells: expect.arrayContaining([expect.objectContaining({ cell: expect.objectContaining({ value: { type: "text", value: "external" } }) })]), errors: [{ code: "RevisionConflict", text: "迟到" }] });
});
test("binding disposal releases DOM listeners and leaves its host session usable", async ({ page }) => {
  await page.getByRole("textbox", { name: "工作簿输入", exact: true }).pressSequentially("draft"); await page.evaluate(() => window.workbookInput.releaseBinding());
  await page.getByRole("textbox", { name: "工作簿输入", exact: true }).press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ revision: 0, history: { undoSteps: 0 }, state: { text: "draft" } });
  await page.evaluate(() => window.workbookInput.dispose()); expect(await page.evaluate(() => window.workbookInput.hostStillUsable())).toBe(true);
});
test("an oversized native draft cannot commit an older controller value and can be corrected", async ({ page }) => {
  const input = page.getByRole("textbox", { name: "工作簿输入", exact: true }); await input.fill("x".repeat(32768));
  expect(await page.evaluate(() => window.workbookInput.paste("must remain a text draft"))).toBe(false); await input.press("Enter");
  expect(await page.evaluate(() => { const state = window.workbookInput.inspect(); return { revision: state.revision, length: state.value.length, errors: [...new Set(state.errors.map(error => error.code))] }; }))
    .toEqual({ revision: 0, length: 32768, errors: ["ResourceLimitExceeded"] });
  await input.fill("corrected"); await input.press("Enter"); expect(await page.evaluate(() => window.workbookInput.inspect())).toMatchObject({ history: { undoSteps: 1 }, state: { text: "corrected", dirty: false } });
});
