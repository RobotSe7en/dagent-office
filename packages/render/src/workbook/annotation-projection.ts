import { DocumentError } from "@doffi/core";
import type { Annotation, DeepReadonly, WorkbookComment, WorkbookRange, WorkbookSnapshot } from "@doffi/core";
import type { AxisDimensions } from "./dimensions.js";
import { intersectRect, paneRect, WorkbookMergeIndex } from "./projection.js";
import type { WorkbookRect } from "./projection.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { checkWindowIdentity } from "./window-source.js";

export interface WorkbookAnnotationRegion {
  readonly rangeIndex: number;
  /** Original source range, even when one cell projects as a merged cell. */
  readonly range: DeepReadonly<WorkbookRange>;
  readonly pane: WorkbookWindowFrame["layout"]["panes"][number]["kind"];
  readonly rect: WorkbookRect;
  readonly clip: WorkbookRect;
}
export interface WorkbookAnnotationReference {
  readonly annotation: DeepReadonly<Annotation>;
  readonly native?: DeepReadonly<WorkbookComment>;
}
export interface WorkbookAnnotationWindow {
  readonly visible: readonly (WorkbookAnnotationReference & { readonly regions: readonly WorkbookAnnotationRegion[] })[];
  /** Keep untrusted locations out of the grid while retaining their full
   * content, provenance and recovery reasons for the owning comment panel. */
  readonly recovery: readonly WorkbookAnnotationReference[];
}
export interface WorkbookAnnotationProjectionLimits {
  readonly maxAnnotations?: number;
  readonly maxRanges?: number;
  readonly maxFragments?: number;
}

/** Source geometry only: no per-cell expansion, Office-cache lookup or status
 * mutation. Status filtering, marker appearance and panel interaction belong
 * to the UI owner. Resolved/closed annotations remain addressable here. */
export function projectWorkbookAnnotations(snapshot: WorkbookSnapshot, frame: WorkbookWindowFrame,
  dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, input: WorkbookAnnotationProjectionLimits = {}): WorkbookAnnotationWindow {
  checkWindowIdentity(frame, snapshot);
  const limits = { maxAnnotations: 5000, maxRanges: 20_000, maxFragments: 20_000, ...input };
  if (Object.keys(input).some(key => !["maxAnnotations", "maxRanges", "maxFragments"].includes(key))
    || !Object.values(limits).every(value => Number.isSafeInteger(value) && value > 0)) throw new DocumentError("InvalidOperation", "Workbook annotation projection requires supported positive budgets.");
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === frame.sheetId);
  if (!sheet) throw new DocumentError("InvalidOperation", "Workbook annotations require their owning worksheet window.");
  if (snapshot.annotations.length > limits.maxAnnotations) throw new DocumentError("ResourceLimitExceeded", "Workbook annotations exceed their source-count budget.");
  const bindings = new Map((snapshot.metadata.workbook.comments ?? []).map(comment => [comment.annotationId, comment]));
  const sheets = new Set(snapshot.metadata.workbook.sheets.map(sheet => sheet.sheetId)), merges = new WorkbookMergeIndex(sheet.merges);
  const visible: (WorkbookAnnotationReference & { readonly regions: readonly WorkbookAnnotationRegion[] })[] = [], recovery: WorkbookAnnotationReference[] = [];
  let ranges = 0, fragments = 0;
  for (const annotation of snapshot.annotations) {
    const anchor = annotation.anchor;
    if (anchor.kind !== "cells") throw new DocumentError("UnsupportedFeature", "This workbook annotation has no cell-range projection.");
    ranges += anchor.ranges.length;
    if (ranges > limits.maxRanges) throw new DocumentError("ResourceLimitExceeded", "Workbook annotations exceed their cumulative range budget.");
    const native = bindings.get(annotation.annotationId), reference = Object.freeze({ annotation, ...(native ? { native } : {}) });
    if (annotation.status === "needs_reanchor") {
      if (anchor.sheetId === frame.sheetId || !sheets.has(anchor.sheetId)) recovery.push(reference);
      continue;
    }
    if (anchor.sheetId !== frame.sheetId) continue;
    const regions: WorkbookAnnotationRegion[] = [];
    for (const [rangeIndex, range] of anchor.ranges.entries()) {
      const single = range.from.row === range.to.row && range.from.column === range.to.column;
      const visual = single ? merges.at(range.from.row, range.from.column) ?? range : range;
      const x = dimensions.columns.position(visual.from.column), y = dimensions.rows.position(visual.from.row);
      const width = dimensions.columns.position(visual.to.column + 1) - x, height = dimensions.rows.position(visual.to.row + 1) - y;
      if (!width || !height) continue;
      for (const pane of frame.layout.panes) {
        const rect = { x: x - pane.sourceLeftPt + pane.inlineStartPt, y: y - pane.sourceTopPt + pane.blockStartPt, width, height };
        const clip = intersectRect(rect, paneRect(pane)); if (!clip) continue;
        if (++fragments > limits.maxFragments) throw new DocumentError("ResourceLimitExceeded", "Workbook annotations exceed their cumulative pane-fragment budget.");
        regions.push(Object.freeze({ rangeIndex, range, pane: pane.kind, rect: Object.freeze(rect), clip: Object.freeze(clip) }));
      }
    }
    if (regions.length) visible.push(Object.freeze({ ...reference, regions: Object.freeze(regions) }));
  }
  return Object.freeze({ visible: Object.freeze(visible), recovery: Object.freeze(recovery) });
}
