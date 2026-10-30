import { fileURLToPath } from "node:url";
import { build } from "vite";

await build({ configFile: false, logLevel: "error", build: {
  emptyOutDir: false,
  outDir: fileURLToPath(new URL("../packages/render/dist/", import.meta.url)), target: "es2022",
  lib: { entry: fileURLToPath(new URL("../packages/render/src/browser-entry.ts", import.meta.url)),
    formats: ["iife"], name: "DagentOfficeRenderer", fileName: () => "browser-script.js" },
} });
