import { DocumentError } from "@doffi/core";
import { svgElement } from "../presentation/svg.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookCellFill } from "./cell-fill.js";
import { createWorkbookFillView } from "./cell-fill-view.js";
import type { WorkbookFillClip } from "./cell-fill-view.js";
import type { WorkbookCellShape } from "./cell-shape.js";
import { workbookCellShapeOffset } from "./cell-shape.js";

/** Inverse-project the visible fragment before creating paint. This retains
 * full-cell gradient/pattern phase without allocating the potentially huge
 * rotated parallelogram. The outer SVG clips at the compositor's pane edge. */
export function createWorkbookShapedFillView(document: Document, paint: WorkbookCellFill, shape: WorkbookCellShape, id: string,
  clip: WorkbookFillClip, maxNodes = 50_000): { readonly element: SVGSVGElement; readonly nodes: number } {
  const { widthPt, heightPt, topOffsetPt } = shape;
  if (![widthPt, heightPt, topOffsetPt, clip.xPt, clip.yPt, clip.widthPt, clip.heightPt].every(Number.isFinite) || widthPt <= 0 || heightPt <= 0
    || clip.widthPt <= 0 || clip.heightPt <= 0 || clip.yPt < 0 || clip.yPt + clip.heightPt > heightPt
    || clip.xPt < Math.min(0, topOffsetPt) || clip.xPt + clip.widthPt > widthPt + Math.max(0, topOffsetPt)
    || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id) || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "A shaped cell fill needs a bounded fragment inside its physical extent.");
  if (![widthPt + Math.abs(topOffsetPt), topOffsetPt / heightPt, clip.xPt + clip.widthPt, clip.yPt + clip.heightPt].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "A shaped cell fill exceeds finite geometry.");
  if (clip.widthPt > 1_000_000 || clip.heightPt > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "A shaped cell fill exceeds its visible extent budget.");
  if (!topOffsetPt) return createWorkbookFillView(document, paint, widthPt, heightPt, id, clip, maxNodes);
  const element = markViewNode(svgElement(document, "svg", { viewBox: `0 0 ${clip.widthPt} ${clip.heightPt}`, "aria-hidden": "true", "data-office-view": "workbook-cell-fill" })) as SVGSVGElement;
  element.style.cssText = `position:absolute;left:0;top:0;width:${clip.widthPt}pt;height:${clip.heightPt}pt;pointer-events:none;overflow:hidden`;
  const offsets = [workbookCellShapeOffset(shape, clip.yPt), workbookCellShapeOffset(shape, clip.yPt + clip.heightPt)];
  const left = Math.max(0, clip.xPt - Math.max(...offsets)), right = Math.min(widthPt, clip.xPt + clip.widthPt - Math.min(...offsets));
  if (right <= left) return Object.freeze({ element, nodes: 1 });
  if (maxNodes <= 2) throw new DocumentError("ResourceLimitExceeded", "A shaped cell fill exceeds its node budget.");
  const source = { xPt: left, yPt: clip.yPt, widthPt: right - left, heightPt: clip.heightPt };
  const fill = createWorkbookFillView(document, paint, widthPt, heightPt, id, source, maxNodes - 2);
  const group = markViewNode(svgElement(document, "g", { transform: `matrix(1 0 ${-topOffsetPt / heightPt} 1 ${left + offsets[0]! - clip.xPt} 0)` }));
  group.append(...fill.element.childNodes); element.append(group);
  // Includes the short-lived inner SVG in the peak node count.
  return Object.freeze({ element, nodes: fill.nodes + 2 });
}
