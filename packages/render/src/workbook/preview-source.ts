import { DocumentError } from "@doffi/core";
import type { WorkbookCalculationResult, WorkbookModel } from "@doffi/core";
import { modelWindowSource } from "./window-source.js";
import { modelAssetSource } from "./assets.js";
import { modelChartSource } from "./chart-source.js";
import { modelRuleSource } from "./rule-source.js";
import type { WorkbookPreviewInput, WorkbookPreviewSource } from "./preview-contracts.js";

/** Direct adapter for a host-owned immutable source and optional completed
 * calculation. Worker hosts create these providers inside their Worker. */
export function createWorkbookPreviewSource(model: WorkbookModel, calculation?: WorkbookCalculationResult): WorkbookPreviewSource {
  return Object.freeze({ kind: "workbook", cells: modelWindowSource(model), assets: modelAssetSource(model), charts: modelChartSource(model, calculation), rules: modelRuleSource(model, calculation) });
}
export function workbookPreviewSource(input: WorkbookPreviewInput): WorkbookPreviewSource {
  if (input.kind !== "workbook") throw new DocumentError("UnsupportedFormat", "Workbook preview requires a workbook model or authorized read provider.");
  return Object.freeze("cells" in input ? { kind: "workbook", cells: input.cells, assets: input.assets, charts: input.charts, ...(input.rules ? { rules: input.rules } : {}) }
    : createWorkbookPreviewSource(input));
}
