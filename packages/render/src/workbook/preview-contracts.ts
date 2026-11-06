import type { DocumentErrorData, DocumentSelection, WorkbookModel, WorkbookSnapshot } from "@doffi/core";
import type { FontResolution } from "../contracts.js";
import type { WorkbookWindowSource } from "./window-source.js";
import type { WorkbookAssetSource } from "./assets.js";
import type { WorkbookChartSource } from "./chart-source.js";
import type { WorkbookRuleSource } from "./rule-source.js";
import type { WorkbookRuleDiagnostic } from "./rule-source.js";
import type { WorkbookDrawingDiagnostic } from "./drawing-layout.js";
import type { WorkbookWindowControllerOptions } from "./window-controller.js";
import type { WorkbookViewport } from "./viewport.js";

/** The host owns these already validated, authorized providers and their
 * Worker/transport lifetime. A preview never serializes the complete workbook
 * or disposes a shared provider. Identifiers alone grant no access. */
export interface WorkbookPreviewSource {
  readonly kind: "workbook";
  readonly cells: WorkbookWindowSource;
  readonly assets: WorkbookAssetSource;
  readonly charts: WorkbookChartSource;
  readonly rules?: WorkbookRuleSource;
}
export type WorkbookPreviewInput = WorkbookModel | WorkbookPreviewSource;
export interface WorkbookPreviewOptions {
  readonly title?: string;
  readonly widthPt?: number;
  readonly heightPt?: number;
  readonly sheetId?: string;
  readonly locale?: string;
  readonly limits?: WorkbookWindowControllerOptions;
  readonly onError?: (error: DocumentErrorData) => void;
  readonly onSelectionChange?: (selection: DocumentSelection | undefined) => void;
}
export interface WorkbookRenderReport {
  readonly kind: "workbook";
  readonly layout: "grid";
  readonly documentId: string;
  readonly revision: number;
  readonly sheetId: string;
  readonly viewport: WorkbookViewport;
  readonly maxScrollLeftPt: number;
  readonly maxScrollTopPt: number;
  readonly visibleCellCount: number;
  /** Includes repeated frozen-pane and offscreen source fragments. */
  readonly cellFragmentCount: number;
  readonly storedCellCount: number;
  readonly objectFragmentCount: number;
  readonly formulaCaches: Readonly<Record<"imported" | "valid" | "stale" | "unavailable", number>>;
  readonly ruleFormulaValues: Readonly<Record<"current" | "unavailable", number>>;
  readonly diagnostics: readonly (WorkbookDrawingDiagnostic | WorkbookRuleDiagnostic)[];
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
}
export interface WorkbookPreview {
  readonly kind: "workbook";
  readonly element: HTMLDivElement;
  readonly ready: Promise<void>;
  /** Metadata and auxiliary content only; cell chunks remain with the provider. */
  snapshot(): WorkbookSnapshot;
  layout(): WorkbookRenderReport;
  fonts(): readonly FontResolution[];
  selection(): DocumentSelection | undefined;
  settled(): Promise<WorkbookRenderReport>;
  update(input: WorkbookPreviewInput): void;
  showSheet(sheetId: string): Promise<WorkbookRenderReport>;
  setViewport(viewport: WorkbookViewport): Promise<WorkbookRenderReport>;
  openComment(annotationId: string): void;
  closeComment(): void;
  dispose(): void;
}
