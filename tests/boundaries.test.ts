import { mkdtempSync, mkdirSync, rmSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { afterEach, expect, test } from "vitest";
import { checkBoundaries } from "../scripts/check-boundaries.mjs";

const temporary: string[] = [];
const layout = ["packages/core", "packages/editor", "packages/office", "packages/node", "apps/playground"];

function fixture(project: string, source: string, dependencies: Record<string, string> = {}) {
  const root = mkdtempSync(join(tmpdir(), "dagent-office-boundaries-"));
  temporary.push(root);
  for (const path of layout) {
    const directory = join(root, path);
    mkdirSync(join(directory, "src"), { recursive: true });
    writeFileSync(join(directory, "src/index.ts"), path === project ? source : "export {};\n");
    writeFileSync(join(directory, "package.json"), JSON.stringify({
      name: "@dagent-office/" + path.split("/").at(-1),
      exports: { ".": { import: "./dist/index.js", types: "./dist/index.d.ts" } },
      dependencies: path === project ? dependencies : {},
    }));
  }
  return root;
}

afterEach(() => {
  for (const path of temporary.splice(0)) rmSync(path, { recursive: true, force: true });
});

test("accepts a public import along the allowed dependency direction", () => {
  const root = fixture("packages/editor", 'export * from "@dagent-office/core";', {
    "@dagent-office/core": "workspace:*",
  });
  expect(checkBoundaries(root)).toEqual([]);
});

test("rejects a reverse dependency even when no source imports it yet", () => {
  const root = fixture("packages/core", "export {};", { "@dagent-office/editor": "workspace:*" });
  expect(checkBoundaries(root)).toEqual(expect.arrayContaining([expect.stringContaining("Forbidden dependency")]));
});

test.each([
  ['export type { Hidden } from "@dagent-office/core/dist/index.js";', { "@dagent-office/core": "workspace:*" }, "public export"],
  ['import "../../core/src/index.js";', {}, "Relative import leaves"],
  ['import "@dagent-office/core";', {}, "Undeclared runtime dependency"],
  ['import "node:fs";', {}, "Node builtin"],
  ['void import("fs/promises");', {}, "Node builtin"],
  ['void import("./" + "hidden.js");', {}, "string literals"],
] as const)("rejects an unsafe browser-module import: %s", (source, dependencies, message) => {
  const root = fixture("packages/editor", source, dependencies);
  expect(checkBoundaries(root).some((value) => value.includes(message))).toBe(true);
});

test("rejects a UI dependency in the portable document core", () => {
  const root = fixture("packages/core", 'import "react";', { react: "19.3.1" });
  expect(checkBoundaries(root).some((value) => value.includes("Host or UI dependency"))).toBe(true);
});
