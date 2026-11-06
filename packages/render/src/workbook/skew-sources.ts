import type { WorkbookSnapshot } from "@doffi/core";
import type { AxisCell, AxisDimensions } from "./dimensions.js";
import type { WorkbookPane, WorkbookViewportLayout } from "./viewport.js";

export interface WorkbookSkewColumns { readonly pane: WorkbookPane["kind"]; readonly columns: readonly AxisCell[] }
/** Custom-property declarations may resolve through variables. Their maximum
 * valid whole-degree cell skew is still bounded by a one-degree rotation.
 * CSSOM reads canonical property names, including escaped source identifiers. */
export function maximumWorkbookSkew(document: Document): number {
  const maximum = 1 / Math.tan(Math.PI / 180); let result = 0;
  const pending = [...document.styleSheets].flatMap(sheet => [...sheet.cssRules]);
  while (pending.length) {
    const rule = pending.pop() as CSSStyleRule & { readonly cssRules?: CSSRuleList };
    if (rule.cssRules) for (const child of rule.cssRules) pending.push(child);
    const value = rule.style?.getPropertyValue("--office-cell-text-rotation").trim(); if (!value) continue;
    const number = /^\d+$/.test(value) ? Number(value) : NaN;
    if (![number >= 0, number <= 180 || number === 255].every(Boolean)) { result = maximum; continue; }
    const angle = number === 255 ? 0 : number <= 90 ? -number : number - 90;
    if (angle && Math.abs(angle) !== 90) result = Math.max(result, Math.abs(1 / Math.tan(angle * Math.PI / 180)));
  }
  return result;
}

/** Conservative source axes for skewed fills and shared borders. The visible
 * grid stays unchanged; only private style/edge discovery uses these columns. */
export function workbookSkewColumns(snapshot: WorkbookSnapshot, sheetId: string, layout: WorkbookViewportLayout,
  dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, coefficient: number, maxCells: number): readonly WorkbookSkewColumns[] {
  if (!coefficient) return Object.freeze([]);
  const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId)!, merges = sheet.merges;
  return Object.freeze(layout.panes.map(pane => {
    let height = 0; for (const row of pane.rows) height = Math.max(height, row.sizePt);
    for (const merge of merges) if (pane.rows.some(row => row.index >= merge.from.row && row.index <= merge.to.row))
      height = Math.max(height, dimensions.rows.position(merge.to.row + 1) - dimensions.rows.position(merge.from.row));
    const extent = height * coefficient;
    const frozen = pane.kind === "corner" || pane.kind === "frozenColumns", boundary = dimensions.columns.position(sheet.frozen.columns);
    const columns = dimensions.columns.visible(Math.max(frozen ? 0 : boundary, pane.sourceLeftPt - extent), Math.min(frozen ? boundary : dimensions.columns.extentPt, pane.sourceLeftPt + pane.widthPt + extent), maxCells);
    for (const column of columns) Object.freeze(column);
    return Object.freeze({ pane: pane.kind, columns: Object.freeze(columns) });
  }));
}
