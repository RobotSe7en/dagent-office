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
const reportSources = Object.fromEntries(
  ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"].map(
    (path) => [path, readFileSync(join(repository, "examples/report", path), "utf8")],
  ),
);
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
  const store = spawnSync(process.execPath, [pnpm, "store", "path"], {
    cwd: repository, encoding: "utf8", timeout: 30_000,
  });
  if (store.error) throw store.error;
  expect(store.status, store.stderr).toBe(0);
  run(process.execPath, [
    pnpm, "install", "--offline", "--ignore-scripts", "--ignore-workspace",
    "--no-frozen-lockfile", "--store-dir", store.stdout.trim(),
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
    'const { CONTRACT_NAMES, DocumentError, getContractJsonSchema, openDocument, parseContract } = await import("@dagent-office/core");',
    "const input = " + readFileSync(join(repository, "examples/report/manifest.json"), "utf8") + ";",
    'assert.equal(parseContract("document-manifest", input).revision, 0);',
    'assert.throws(() => parseContract("document-manifest", { ...input, formatVersion: 2 }),',
    '  (error) => error instanceof DocumentError && error.code === "UnsupportedVersion");',
    "const sources = " + JSON.stringify(reportSources) + ";",
    "const files = new Map(Object.entries(sources).map(([path, source]) => [path, new TextEncoder().encode(source)]));",
    "const document = openDocument(files);",
    'assert.equal(document.getNode("n-summary").text, "项目计划于第三季度完成。");',
    "assert.deepEqual(openDocument(document.serialize()).snapshot(), document.snapshot());",
    'const { createAssetDescriptor, collectDocumentAssets } = await import("@dagent-office/core");',
    'const { readDocumentZip, writeDocumentZip } = await import("@dagent-office/office");',
    'assert.equal(createAssetDescriptor("assets/a.bin", "application/octet-stream", new Uint8Array([97,98,99])).sha256, "ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad");',
    'const collected = await collectDocumentAssets(files, { async readAsset() { throw new Error("Unexpected host read"); } });',
    'assert.deepEqual(openDocument(readDocumentZip(writeDocumentZip(collected))).snapshot(), document.snapshot());',
    "for (const name of CONTRACT_NAMES) {",
    '  const schema = await import("@dagent-office/core/schemas/" + name + ".json", { with: { type: "json" } });',
    "  assert.deepEqual(schema.default, getContractJsonSchema(name));",
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
    ["core", ["core", "office"], ["ES2022"]],
    ["browser", ["core", "editor", "office"], ["ES2022", "DOM", "DOM.Iterable"]],
  ] as const) {
    writeFileSync(join(consumer, name + ".ts"), [...imported.flatMap((module, index) => [
      "import * as Module" + index + ' from "@dagent-office/' + module + '";',
      "void Module" + index + ";",
      "// @ts-expect-error Internal paths are outside the public package contract.",
      "type Private" + index + ' = typeof import("@dagent-office/' + module + '/dist/index.js");',
    ]), 'import { openDocument, parseContract } from "@dagent-office/core";',
      'const manifest = parseContract("document-manifest", {});',
      "const version: 1 = manifest.formatVersion;",
      "void version;",
      "// @ts-expect-error The registry must preserve the selected output type.",
      "manifest.proposalId;",
      "declare const files: ReadonlyMap<string, Uint8Array>;",
      "const document = openDocument(files);",
      'const text: string | undefined = document.getNode("n")?.text;',
      "void text;",
      "// @ts-expect-error Parser-specific AST is not part of the public node interface.",
      'document.getNode("n")?.childNodes;',
      'import { collectDocumentAssets, createAssetDescriptor } from "@dagent-office/core";',
      'import type { AssetProvider } from "@dagent-office/core";',
      'import { readDocumentZip, writeDocumentZip } from "@dagent-office/office";',
      'const provider: AssetProvider = { async readAsset() { return undefined; } };',
      'const collected: Promise<ReadonlyMap<string, Uint8Array>> = collectDocumentAssets(files, provider);',
      'const archive: Uint8Array = writeDocumentZip(files);',
      'const unpacked: ReadonlyMap<string, Uint8Array> = readDocumentZip(archive);',
      'const digest: string = createAssetDescriptor("assets/a.bin", "application/octet-stream", archive).sha256;',
      'void collected; void unpacked; void digest;',
    ].join("\n"));
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
  ).join("\n") + '\nimport { openDocument, parseContract } from "@dagent-office/core";\n'
    + "const input = " + readFileSync(join(repository, "examples/report/manifest.json"), "utf8") + ";\n"
    + 'document.documentElement.dataset.revision = String(parseContract("document-manifest", input).revision);\n'
    + "const sources = " + JSON.stringify(reportSources) + ";\n"
    + "const files = new Map(Object.entries(sources).map(([path, source]) => [path, new TextEncoder().encode(source)]));\n"
    + 'document.documentElement.dataset.documentId = openDocument(files).snapshot().manifest.documentId;\n'
    + 'import { createAssetDescriptor, collectDocumentAssets } from "@dagent-office/core";\n'
    + 'import { readDocumentZip, writeDocumentZip } from "@dagent-office/office";\n'
    + 'const collected = await collectDocumentAssets(files, { async readAsset() { return undefined; } });\n'
    + 'document.documentElement.dataset.roundtrip = openDocument(readDocumentZip(writeDocumentZip(collected))).snapshot().manifest.documentId;\n'
    + 'document.documentElement.dataset.digest = createAssetDescriptor("assets/a.bin", "application/octet-stream", new Uint8Array([97,98,99])).sha256;\n');
  await build({
    root: consumer,
    configFile: false,
    logLevel: "silent",
    build: { target: "es2022", outDir: "browser-dist" },
  });
  expect(existsSync(join(consumer, "browser-dist/index.html"))).toBe(true);
  expect(existsSync(join(consumer, "node_modules/react"))).toBe(false);
});
