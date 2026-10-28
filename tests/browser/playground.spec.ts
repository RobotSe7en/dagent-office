import { expect, test } from "@playwright/test";
import { readFileSync } from "node:fs";
import { fileURLToPath } from "node:url";

const samples = fileURLToPath(new URL("../../examples/docx/", import.meta.url));
const cases = JSON.parse(
  readFileSync(
    new URL("../../examples/docx/expected.json", import.meta.url),
    "utf8",
  ),
) as { file: string; requiredText: string[] }[];

test("opens a DOCX sample through the public APIs and releases its preview on replacement", async ({
  page,
}) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  await page.goto("/");
  await expect(
    page.getByRole("heading", { name: "dagent-office Playground" }),
  ).toBeVisible();
  await page.getByRole("button", { name: "打开 DOCX 样例" }).click();
  await expect(page.getByRole("status")).toContainText("已打开");
  const frame = page.frameLocator('iframe[title="文档只读预览"]');
  await expect(frame.getByRole("heading", { name: "实施报告" })).toBeVisible();
  await expect(frame.locator(".sdk-number-marker")).toHaveText([
    "1.",
    "1.a)",
    "2.",
    "4.",
  ]);
  const color = await frame
    .getByRole("heading", { name: "实施报告" })
    .evaluate((node) => getComputedStyle(node).color);
  expect(color).toBe("rgb(31, 75, 153)");
  const inherited = await frame
    .getByRole("heading", { name: "实施报告" })
    .evaluate((node) => {
      (node as HTMLElement).style.color = "#884422";
      return getComputedStyle(node.querySelector("[data-word-run]")!).color;
    });
  expect(inherited).toBe("rgb(136, 68, 34)");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(frame.locator("ins")).toHaveText("新安排");
  await expect(frame.locator("del")).toHaveText("旧安排");
  await expect(frame.getByText("请修改", { exact: false })).toBeVisible();
  await expect(page.locator("iframe")).toHaveCount(1);
  expect(errors).toEqual([]);
});
for (const item of cases) {
  test("read-only visual fixture: " + item.file, async ({ page }) => {
    const external: string[] = [];
    page.on("request", (request) => {
      if (request.url().startsWith("https://example.com"))
        external.push(request.url());
    });
    await page.goto("/");
    await page.getByLabel("打开文件").setInputFiles(samples + item.file);
    await expect(page.getByRole("status")).toContainText("已打开");
    const frame = page.frameLocator("iframe");
    await expect(frame.locator("article")).toBeVisible();
    await frame.locator("body").evaluate(async () => {
      await document.fonts.ready;
      await Promise.all([...document.images].map((image) => image.decode()));
    });
    if (item.file === "03-images.docx")
      await expect(frame.locator("img")).toHaveCount(3);
    if (item.file === "06-links-tabs.docx") {
      const positions = await frame
        .locator("[data-word-tab]")
        .evaluate((tab) => ({
          left: tab.getBoundingClientRect().left,
          right: tab.getBoundingClientRect().right,
        }));
      expect(positions.right).toBeGreaterThan(positions.left);
      await expect(frame.locator("a[href]")).toHaveCount(0);
    }
    await expect(frame.locator("article")).toHaveScreenshot(
      item.file.replace(".docx", ".png"),
      { animations: "disabled", maxDiffPixelRatio: 0.005, threshold: 0.2 },
    );
    expect(external).toEqual([]);
  });
}
test("downloads its document package and reopens it", async ({
  page,
}, info) => {
  await page.goto("/");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(page.getByRole("status")).toContainText("已打开");
  const download = page.waitForEvent("download");
  await page.getByRole("button", { name: "下载文档包" }).click();
  const path = info.outputPath("document.zip");
  await (await download).saveAs(path);
  await page.getByLabel("打开文件").setInputFiles(path);
  await expect(page.getByRole("status")).toContainText("批注 1");
  await expect(page.frameLocator("iframe").locator("ins")).toHaveText("新安排");
});

test("keeps distinct Latin and Chinese faces and reports missing Chinese fonts", async ({
  page,
}) => {
  await page.goto("/");
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "09-chinese-fonts.docx");
  await expect(page.getByRole("status")).toContainText("已打开");
  const frame = page.frameLocator("iframe");
  const source = frame.locator('[data-word-script="eastAsia"]').first();
  await expect(source).toBeVisible();
  expect(
    await source.evaluate((node) => getComputedStyle(node).fontFamily),
  ).toContain("Noto Serif CJK SC");
  await expect(frame.getByRole("note")).toContainText(
    "Dagent Office Missing CJK Test Font",
  );
  await expect(frame.getByRole("note")).toContainText("U+20000");
  const sourceNames = await frame
    .locator("article")
    .evaluate((node) => node.outerHTML);
  expect(sourceNames).toContain("--word-font-eastAsia");
});

test("releases image URLs when switching documents", async ({ page }) => {
  await page.addInitScript(() => {
    const counts = { created: 0, revoked: 0 };
    (window as unknown as { urlCounts: typeof counts }).urlCounts = counts;
    const create = URL.createObjectURL.bind(URL),
      revoke = URL.revokeObjectURL.bind(URL);
    URL.createObjectURL = (value) => {
      counts.created++;
      return create(value);
    };
    URL.revokeObjectURL = (value) => {
      counts.revoked++;
      revoke(value);
    };
  });
  await page.goto("/");
  await page.getByLabel("打开文件").setInputFiles(samples + "03-images.docx");
  await expect(page.frameLocator("iframe").locator("img")).toHaveCount(3);
  await page
    .getByLabel("打开文件")
    .setInputFiles(samples + "08-comments-revisions.docx");
  await expect(page.frameLocator("iframe").locator("ins")).toHaveText("新安排");
  expect(
    await page.evaluate(
      () =>
        (
          window as unknown as {
            urlCounts: { created: number; revoked: number };
          }
        ).urlCounts,
    ),
  ).toEqual({ created: 3, revoked: 3 });
});
