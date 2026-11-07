import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import { openSdkHost } from "./sdk-host.js";
import type * as fixture from "./fixtures/workbook-input.js";
declare global { interface Window { workbookInput: typeof fixture } }
let code = "";
test.beforeAll(async () => {
  const result = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: { entry: fileURLToPath(new URL("./fixtures/workbook-input.ts", import.meta.url)), formats: ["iife"], name: "WorkbookInput" } } });
  const output = Array.isArray(result) ? result.flatMap(part => part.output) : "output" in result ? result.output : [];
  code = output.flatMap(part => part.type === "chunk" ? [part.code] : []).join("\n");
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
