import { fileURLToPath } from "node:url";
import { build } from "vite";
import { readFileSync } from "node:fs";

const notice = readFileSync(new URL("../packages/core/ECMA-376-GEOMETRY-LICENSE.txt", import.meta.url), "utf8");

await build({ configFile: false, logLevel: "error", build: {
  emptyOutDir: false,
  rolldownOptions: { output: { banner: "/*! @license\n" + notice + "\n*/", comments: { legal: true } } },
  outDir: fileURLToPath(new URL("../packages/render/dist/", import.meta.url)), target: "es2022",
  lib: { entry: fileURLToPath(new URL("../packages/render/src/browser-entry.ts", import.meta.url)),
    formats: ["iife"], name: "DagentOfficeRenderer", fileName: () => "browser-script.js" },
} });
