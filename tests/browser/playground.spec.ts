import { expect, test } from "@playwright/test";

test("loads the built React playground without browser errors", async ({ page }) => {
  const errors: string[] = [];
  page.on("pageerror", (error) => errors.push(error.message));
  page.on("console", (message) => {
    if (message.type() === "error") errors.push(message.text());
  });

  await page.goto("/");
  await expect(page).toHaveTitle("dagent-office Playground");
  await expect(page.getByRole("heading", { name: "dagent-office Playground" })).toBeVisible();
  await expect(page.getByText("工程环境已就绪。")).toBeVisible();
  await expect(page.getByText("文档读取、编辑与协作功能将在后续任务中接入。")).toBeVisible();
  expect(errors).toEqual([]);
});
