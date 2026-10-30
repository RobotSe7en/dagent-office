import { fileURLToPath } from "node:url";
import { readFileSync } from "node:fs";
import react from "@vitejs/plugin-react";
import { defineConfig } from "vite";
import type { Plugin } from "vite";

/** Public synthetic inputs belong to this development host, never to SDK runtime I/O. */
function collaborationExamples(): Plugin {
  const files = ["manifest.json", "content.html", "styles.css", "metadata.json", "annotations.json"];
  const assets: Record<string, string> = {
    "report-demo.json": JSON.stringify(files.map((path) => [path, readFileSync(new URL("../../examples/report/" + path, import.meta.url), "utf8")])),
    "report.proposal.json": readFileSync(new URL("../../examples/report.proposal.json", import.meta.url), "utf8"),
  };
  return {
    name: "collaboration-examples",
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        const source = assets[(request.url ?? "").split("?")[0]!.replace(/^\//, "")];
        if (source === undefined) { next(); return; }
        response.setHeader("Content-Type", "application/json; charset=utf-8"); response.end(source);
      });
    },
    generateBundle() {
      for (const [fileName, source] of Object.entries(assets)) this.emitFile({ type: "asset", fileName, source });
    },
  };
}

export default defineConfig({
  plugins: [react(), collaborationExamples()],
  publicDir: fileURLToPath(new URL("../../examples/docx/", import.meta.url)),
  build: { target: "es2022" },
});
