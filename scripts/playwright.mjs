import { spawnSync } from "node:child_process";
import { createRequire } from "node:module";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
const require = createRequire(import.meta.url);
const result = spawnSync(process.execPath, [
  require.resolve("@playwright/test/cli"), ...process.argv.slice(2),
], {
  cwd: root,
  stdio: "inherit",
  env: {
    ...process.env,
    PLAYWRIGHT_BROWSERS_PATH: process.env.PLAYWRIGHT_BROWSERS_PATH ?? resolve(root, ".playwright"),
    PLAYWRIGHT_SKIP_BROWSER_GC: "1",
  },
});
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
