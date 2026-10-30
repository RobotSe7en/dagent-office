import { defineConfig } from "@playwright/test";

export default defineConfig({
  testDir: "./tests/browser",
  fullyParallel: false,
  workers: 1,
  forbidOnly: Boolean(process.env.CI),
  retries: 0,
  snapshotPathTemplate: "{testDir}/{testFilePath}-snapshots/{arg}-{platform}{ext}",
  projects: [
    { name: "chromium", use: { browserName: "chromium", ...(process.env.DAGENT_OFFICE_BROWSER_EXECUTABLE ? { launchOptions: { executablePath: process.env.DAGENT_OFFICE_BROWSER_EXECUTABLE } } : {}) } },
    { name: "firefox", testMatch: /(?:pagination|compatibility|remote-access)\.spec\.ts/, use: { browserName: "firefox" } },
    { name: "webkit", testMatch: /(?:pagination|compatibility|remote-access)\.spec\.ts/, use: { browserName: "webkit" } },
  ],
  use: {
    browserName: "chromium",
    ...(process.env.DAGENT_OFFICE_BROWSER_CHANNEL
      ? { channel: process.env.DAGENT_OFFICE_BROWSER_CHANNEL }
      : {}),
    baseURL: "http://127.0.0.1:4173",
    viewport: { width: 1280, height: 800 },
    trace: "retain-on-failure",
  },
  webServer: {
    command: "pnpm --filter @dagent-office/playground preview --port 4173 --strictPort",
    url: "http://127.0.0.1:4173",
    reuseExistingServer: false,
    timeout: 30_000,
  },
});
