import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { build } from "vite";
import { expect, test } from "@playwright/test";
import type { Page } from "@playwright/test";
import { openDocument } from "@dagent-office/core";
import { importPptx, readDocumentZip } from "@dagent-office/office";
import type {} from "./fixtures/interaction-host.js";

async function report(page: Page): Promise<void> {
  await page.goto("/");
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toContainText("修订 0");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toBeVisible();
}
async function selectText(page: Page, from = 2, to = 8): Promise<void> {
  await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]').evaluate((node, { from, to }) => {
    const range = document.createRange(); range.setStart(node.firstChild!, from); range.setEnd(node.firstChild!, to);
    const selected = document.getSelection()!; selected.removeAllRanges(); selected.addRange(range);
    document.dispatchEvent(new Event("selectionchange"));
  }, { from, to });
}

test("document library filters and switches between isolated sessions while retaining history", async ({ page }) => {
  await report(page);
  const sidebar = page.getByRole("complementary", { name: "文档列表", exact: true });
  expect((await sidebar.boundingBox())!.x).toBeLessThan((await page.getByRole("region", { name: "文档工作区", exact: true }).boundingBox())!.x);
  await page.getByRole("button", { name: "载入固定示例提案", exact: true }).click();
  await page.getByRole("button", { name: "整批应用提案", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toContainText("修订 1");
  await page.getByRole("group", { name: "筛选文档" }).getByRole("button", { name: "PPT", exact: true }).click();
  await expect(sidebar.getByRole("button", { name: "打开报告协作样例", exact: true })).toHaveCount(0);
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "已打开" })).toContainText("修订 0");
  await page.getByRole("group", { name: "筛选文档" }).getByRole("button", { name: "全部", exact: true }).click();
  await page.getByLabel("搜索文档", { exact: true }).fill("季度");
  await expect(sidebar.getByRole("navigation").getByRole("button")).toHaveCount(1);
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toContainText("修订 1");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toContainText("修订 2");
  await expect(page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]')).toContainText("第三季度");
});

test("selection opens a bounded comment dialog; cancellation, IME and keyboard submission preserve transaction semantics", async ({ page }, info) => {
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toBeVisible(); await expect(dialog).toContainText("计划于第三季");
  const input = dialog.getByLabel("新批注内容", { exact: true });
  await input.fill("取消的草稿"); await input.press("Escape");
  await expect(dialog).toBeHidden(); await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 0 · 批注 2");
  await selectText(page); await expect(input).toHaveValue(""); await input.fill("补充时间与负责人");
  await input.evaluate(node => node.dispatchEvent(new KeyboardEvent("keydown", { key: "Enter", ctrlKey: true, isComposing: true, bubbles: true, cancelable: true })));
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 0 · 批注 2");
  const box = (await dialog.boundingBox())!, viewport = page.viewportSize()!;
  expect(box.x).toBeGreaterThanOrEqual(0); expect(box.y).toBeGreaterThanOrEqual(0);
  const toolbar = (await page.getByRole("toolbar", { name: "文档交互", exact: true }).boundingBox())!;
  expect(box.y).toBeGreaterThanOrEqual(toolbar.y + toolbar.height);
  expect(box.x + box.width).toBeLessThanOrEqual(viewport.width); expect(box.y + box.height).toBeLessThanOrEqual(viewport.height);
  await page.screenshot({ path: info.outputPath("word-comment.png") });
  await input.press("Control+Enter"); await expect(dialog).toBeHidden();
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 1 · 批注 3");
  const pending = page.waitForEvent("download"); await page.getByRole("button", { name: "下载文档包", exact: true }).click();
  const path = info.outputPath("comment.zip"); await (await pending).saveAs(path);
  const model = openDocument(readDocumentZip(new Uint8Array(readFileSync(path))));
  expect(model.snapshot().annotations.at(-1)).toMatchObject({ instruction: "补充时间与负责人", anchor: { kind: "text", nodeId: "n-summary", from: 2, to: 8, quote: "计划于第三季" } });
  await selectText(page); await input.fill("过期草稿");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(dialog).toBeHidden(); await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 2 · 批注 2");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 3 · 批注 3");
});

test("object and rectangle selections use the same floating composer", async ({ page }) => {
  await report(page);
  await page.getByRole("button", { name: "对象选择", exact: true }).click();
  await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-title"]').click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toContainText("对象：项目实施报告");
  const input = dialog.getByLabel("新批注内容");
  await expect(input).not.toBeFocused();
  // A non-modal composer opens without focus; enter it as a user would. Firefox
  // keeps the iframe's native text selection until this explicit focus transfer.
  await input.click(); await expect(input).toBeFocused(); await input.fill("对象批注");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 1 · 批注 3");
  await page.getByRole("button", { name: "矩形框选", exact: true }).click();
  const paragraph = await page.frameLocator('iframe[title="文档只读预览"]').locator('[data-node-id="n-summary"]').boundingBox();
  await page.mouse.move(paragraph!.x + 1, paragraph!.y + 1); await page.mouse.down();
  await page.mouse.move(paragraph!.x + 100, paragraph!.y + 20, { steps: 4 });
  await expect(dialog).toBeHidden();
  await page.mouse.up();
  await expect(dialog).toBeVisible(); await input.click(); await input.fill("区域批注");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 2 · 批注 4");
});

test("PPT comment popover preserves native comments and can be undone, exported and reopened", async ({ page }, info) => {
  await page.setViewportSize({ width: 1600, height: 1000 }); await page.goto("/");
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  const frame = page.frameLocator('iframe[title="幻灯片只读预览"]');
  await frame.locator('[data-node-id="heading-p"]').click();
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true }); await expect(dialog).toBeVisible();
  await dialog.getByLabel("批注显示署名", { exact: true }).fill("审阅人");
  await dialog.getByLabel("新批注正文", { exact: true }).fill("确认演示标题");
  await dialog.getByRole("button", { name: "添加批注", exact: true }).click(); await expect(dialog).toBeHidden();
  await expect(page.getByRole("status").filter({ hasText: "已打开" })).toHaveText("已打开 · 修订 1 · 批注 2");
  await page.getByRole("button", { name: "撤销", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "已打开" })).toHaveText("已打开 · 修订 2 · 批注 1");
  await page.getByRole("button", { name: "重做", exact: true }).click();
  const pending = page.waitForEvent("download"); await page.getByRole("button", { name: "导出 PPTX", exact: true }).click();
  const path = info.outputPath("comment.pptx"); await (await pending).saveAs(path);
  const model = openDocument(importPptx(new Uint8Array(readFileSync(path)), { documentId: "workspace-comment" }));
  expect(model.snapshot().annotations).toHaveLength(2);
  expect(model.snapshot().annotations.map(annotation => annotation.bodyNodeId ? model.getNode(annotation.bodyNodeId)?.text : annotation.instruction)).toContain("确认演示标题");
  expect(model.snapshot().annotations.at(-1)?.author?.name).toBe("审阅人");
});

test("switching documents discards comment drafts without changing either session", async ({ page }) => {
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await dialog.getByLabel("新批注内容", { exact: true }).fill("不能带到另一文档");
  await page.getByRole("button", { name: "打开 PPT 协作样例", exact: true }).click();
  await expect(page.getByRole("status").filter({ hasText: "已打开" })).toHaveText("已打开 · 修订 0 · 批注 1");
  await expect(dialog).toBeHidden();
  await page.getByRole("button", { name: "打开报告协作样例", exact: true }).click();
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 0 · 批注 2");
  await selectText(page);
  await expect(dialog.getByLabel("新批注内容", { exact: true })).toHaveValue("");
});

test("comment popover follows scrolling inside the canvas shadow tree", async ({ page }) => {
  await page.setViewportSize({ width: 1600, height: 1000 });
  await report(page); await selectText(page);
  const dialog = page.getByRole("dialog", { name: "新建批注", exact: true });
  await expect(dialog).toBeVisible();
  const before = (await dialog.boundingBox())!;
  const scrolled = await page.locator(".stage").evaluate(node => { node.scrollTop = 80; return node.scrollTop; });
  expect(scrolled).toBeGreaterThan(0);
  await expect.poll(async () => Math.abs((await dialog.boundingBox())!.y - before.y)).toBeGreaterThan(10);
  await expect(page.getByLabel("会话状态", { exact: true })).toHaveText("修订 0 · 批注 2");
});

test("workspace unmount releases views and floating controls while the host retains its session", async ({ page }) => {
  const built = await build({ configFile: false, logLevel: "silent", build: { write: false, target: "es2022", lib: {
    entry: fileURLToPath(new URL("./fixtures/interaction-host.ts", import.meta.url)), formats: ["es"], fileName: "interaction-host",
  } } });
  if (!("output" in built) && !Array.isArray(built)) throw new Error("Missing fixture output");
  const output = Array.isArray(built) ? built.flatMap(item => item.output) : built.output;
  await page.goto("/");
  await page.addScriptTag({ type: "module", content: output.flatMap(item => item.type === "chunk" ? [item.code] : []).join("\n") });
  await page.waitForFunction(() => Boolean(window.interactionHost));
  const files = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(path => [path, [...readFileSync(new URL("../../examples/report/" + path, import.meta.url))]] as [string, number[]]);
  for (let iteration = 0; iteration < 3; iteration++) {
    await page.evaluate(files => window.interactionHost.mount(files, false, "workspace"), files);
    await selectText(page);
    await expect(page.getByRole("dialog", { name: "新建批注", exact: true })).toBeVisible();
    await page.evaluate(() => { window.interactionHost.view!.dispose(); window.interactionHost.view!.dispose(); });
    await expect(page.locator("iframe")).toHaveCount(0);
    await expect(page.getByRole("dialog", { name: "新建批注", exact: true })).toHaveCount(0);
    expect(await page.evaluate(() => ({ revision: window.interactionHost.session!.snapshot().manifest.revision, text: window.interactionHost.session!.getNode("n-summary")!.text, errors: window.interactionHost.errors }))).toMatchObject({ revision: 0, errors: [], text: expect.stringContaining("第三季度") });
  }
  await page.evaluate(() => window.interactionHost.session!.dispose());
});

test("narrow workspace keeps page chrome within the viewport and comment actions above the canvas", async ({ page }, info) => {
  await page.setViewportSize({ width: 390, height: 844 }); await report(page);
  expect(await page.evaluate(() => document.documentElement.scrollWidth)).toBeLessThanOrEqual(390);
  const toolbar = (await page.getByRole("toolbar", { name: "文件与导出", exact: true }).boundingBox())!;
  const canvas = (await page.locator('iframe[title="文档只读预览"]').boundingBox())!;
  expect(toolbar.y + toolbar.height).toBeLessThan(canvas.y);
  await selectText(page); const dialog = page.getByRole("dialog", { name: "新建批注", exact: true }); await expect(dialog).toBeVisible();
  const box = (await dialog.boundingBox())!; expect(box.x + box.width).toBeLessThanOrEqual(390); expect(box.y + box.height).toBeLessThanOrEqual(844);
  await page.screenshot({ path: info.outputPath("mobile-workspace.png"), fullPage: true });
});
