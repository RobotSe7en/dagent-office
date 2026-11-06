import { DocumentError } from "@dagent-office/core";
import { modelWindowSource } from "./window-source.js";
import { modelAssetSource } from "./assets.js";
import { modelChartSource } from "./chart-source.js";
import { modelRuleSource } from "./rule-source.js";
import type { WorkbookPreviewInput, WorkbookPreviewSource } from "./preview-contracts.js";

export function workbookPreviewSource(input: WorkbookPreviewInput): WorkbookPreviewSource {
  if (input.kind !== "workbook") throw new DocumentError("UnsupportedFormat", "Workbook preview requires a workbook model or authorized read provider.");
  return Object.freeze("cells" in input ? { kind: "workbook", cells: input.cells, assets: input.assets, charts: input.charts, ...(input.rules ? { rules: input.rules } : {}) }
    : { kind: "workbook", cells: modelWindowSource(input), assets: modelAssetSource(input), charts: modelChartSource(input), rules: modelRuleSource(input) });
}
