import { indexWorkbookChunk } from "../../../packages/core/src/workbook/chunk-index.js";
import type { ChunkIndex } from "../../../packages/core/src/workbook/chunk-index.js";
import { createFormulaTemplatePool } from "../../../packages/core/src/workbook/formula.js";
import { baselineRows, generateBlock } from "../../../scripts/workbook-baseline-fixture.mjs";

// Private storage probe only. No editor, XLSX importer or SDK session is exposed.
let generation = 0, current: ChunkIndex[] = [];
self.addEventListener("message", async (event: MessageEvent<{ kind: "open" | "cancel" | "window"; id: number; row?: number }>) => {
  const { kind, id } = event.data;
  if (kind === "cancel") { generation++; self.postMessage({ kind: "cancelled", id, retainedChunks: current.length }); return; }
  if (kind === "window") {
    const row = event.data.row!;
    self.postMessage({ kind: "window", id, cells: current[Math.floor(row / 256)]?.readWindow({ row, column: 0, rows: 1, columns: 20 }) ?? [] });
    return;
  }
  const token = ++generation, candidate: ChunkIndex[] = [], start = performance.now(), formulas = createFormulaTemplatePool();
  try {
    for (let row = 0; row < baselineRows; row += 256) {
      candidate.push(indexWorkbookChunk(generateBlock(row), { sheetId: "baseline", row, column: 0 }, undefined, formulas));
      self.postMessage({ kind: "progress", id, rows: Math.min(row + 256, baselineRows) });
      await new Promise(resolve => setTimeout(resolve, 0));
      if (generation !== token) return;
    }
    current = candidate;
    self.postMessage({ kind: "opened", id, rows: baselineRows, cellCount: current.reduce((sum, c) => sum + c.cellCount, 0),
      formulaCount: current.reduce((sum, c) => sum + c.formulaCount, 0), formulaTemplates: formulas.size, indexBytes: current.reduce((sum, c) => sum + c.indexBytes, 0), elapsedMs: performance.now() - start });
  } catch (error) {
    if (generation === token) self.postMessage({ kind: "error", id, message: error instanceof Error ? error.message : "Storage probe failed" });
  }
});
