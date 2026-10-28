import { spawnSync } from "node:child_process";
import { existsSync, mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from "node:fs";
import { createRequire } from "node:module";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { fileURLToPath } from "node:url";
import { afterAll, beforeAll, expect, test } from "vitest";
import { build } from "vite";
import { packPackages } from "../scripts/pack.mjs";

const repository = fileURLToPath(new URL("../", import.meta.url));
const require = createRequire(import.meta.url);
const modules = ["core", "editor", "office", "node"];
let temporary: string;
let consumer: string;

function run(executable: string, args: string[], cwd: string) {
  const result = spawnSync(executable, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, CI: "true" },
    timeout: 60_000,
  });
  if (result.error) throw result.error;
  expect(result.status, result.stdout + result.stderr).toBe(0);
}

beforeAll(() => {
  temporary = mkdtempSync(join(tmpdir(), "dagent office consumer-"));
  consumer = join(temporary, "consumer");
  mkdirSync(consumer);
  const archives = packPackages(join(temporary, "archives"));
  expect(archives).toHaveLength(4);
  const dependencies: Record<string, string> = {};
  for (const module of modules) {
    const manifest = JSON.parse(readFileSync(join(repository, "packages", module, "package.json"), "utf8"));
    const filename = manifest.name.replace("@", "").replace("/", "-") + "-" + manifest.version + ".tgz";
    const archive = archives.find((path) => path.endsWith(filename));
    expect(archive).toBeDefined();
    dependencies[manifest.name] = "file:" + archive;
  }
  writeFileSync(join(consumer, "package.json"), JSON.stringify({
    name: "external-consumer",
    private: true,
    type: "module",
    dependencies,
    pnpm: { overrides: dependencies },
  }, null, 2));
  const pnpm = process.env.npm_execpath;
  if (!pnpm) throw new Error("Run package-consumption tests through pnpm.");
  run(process.execPath, [
    pnpm, "install", "--offline", "--ignore-scripts", "--ignore-workspace", "--no-frozen-lockfile",
  ], consumer);
});

afterAll(() => {
  if (temporary) rmSync(temporary, { recursive: true, force: true });
});

test("packed packages load in Node and reject access to their internal files", () => {
  writeFileSync(join(consumer, "node.mjs"), [
    'import assert from "node:assert/strict";',
    "for (const name of " + JSON.stringify(modules.map((name) => "@dagent-office/" + name)) + ") {",
    "  const module = await import(name);",
    '  assert.equal(Object.prototype.toString.call(module), "[object Module]");',
    '  await assert.rejects(import(name + "/dist/index.js"), { code: "ERR_PACKAGE_PATH_NOT_EXPORTED" });',
    "}",
  ].join("\n"));
  run(process.execPath, ["node.mjs"], consumer);
  for (const module of modules) {
    const manifest = JSON.parse(readFileSync(join(consumer, "node_modules/@dagent-office", module, "package.json"), "utf8"));
    for (const version of Object.values(manifest.dependencies ?? {})) {
      expect(version).not.toMatch(/^workspace:/);
    }
  }
});

test("packed declarations resolve without borrowing workspace source or ambient types", () => {
  const compiler = require.resolve("typescript/bin/tsc");
  for (const [name, imported, lib] of [
    ["core", ["core"], ["ES2022"]],
    ["browser", ["core", "editor", "office"], ["ES2022", "DOM", "DOM.Iterable"]],
  ] as const) {
    writeFileSync(join(consumer, name + ".ts"), imported.flatMap((module, index) => [
      "import * as Module" + index + ' from "@dagent-office/' + module + '";',
      "void Module" + index + ";",
      "// @ts-expect-error Internal paths are outside the public package contract.",
      "type Private" + index + ' = typeof import("@dagent-office/' + module + '/dist/index.js");',
    ]).join("\n"));
    writeFileSync(join(consumer, "tsconfig." + name + ".json"), JSON.stringify({
      compilerOptions: {
        target: "ES2022", module: "NodeNext", moduleResolution: "NodeNext",
        strict: true, noEmit: true, types: [], lib,
      },
      files: [name + ".ts"],
    }));
    run(process.execPath, [compiler, "--project", "tsconfig." + name + ".json"], consumer);
  }
});

test("packed browser modules build in a plain JavaScript host without React", async () => {
  writeFileSync(join(consumer, "index.html"), '<main>External browser host</main><script type="module" src="/browser.js"></script>');
  writeFileSync(join(consumer, "browser.js"), ["core", "editor", "office"].map(
    (module) => 'import "@dagent-office/' + module + '";',
  ).join("\n"));
  await build({
    root: consumer,
    configFile: false,
    logLevel: "silent",
    build: { target: "es2022", outDir: "browser-dist" },
  });
  expect(existsSync(join(consumer, "browser-dist/index.html"))).toBe(true);
  expect(existsSync(join(consumer, "node_modules/react"))).toBe(false);
});
