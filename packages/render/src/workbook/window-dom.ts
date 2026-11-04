import type { WorkbookCell } from "@dagent-office/core";
import type { WorkbookGridCell } from "./projection.js";
import type { WorkbookWindowFrame } from "./window-loader.js";
import { markViewNode } from "../view-nodes.js";

/** Paint content/style into a full-size cell body. The surrounding clip retains
 * its pane position; content must not be re-centered inside a visible fragment. */
export type WorkbookCellPainter = (body: HTMLElement, source: WorkbookCell | undefined, projection: WorkbookGridCell) => void;

/** Internal virtual DOM compositor, not an XLSX preview entry point. The source
 * painter owns number formats, rich text and appearance. Geometry and source
 * identity remain independent of those concerns and of physical RTL positions. */
export function replaceWorkbookWindow(container: HTMLElement, frame: WorkbookWindowFrame, paint: WorkbookCellPainter, rightToLeft: boolean, borders: readonly SVGSVGElement[] = []): HTMLElement {
  const document = container.ownerDocument, root = markViewNode(document.createElement("div")), rows = new Map<number, HTMLElement>(), accessible = new Set<number>();
  root.dataset.officeView = "workbook-window"; root.dataset.sheetId = frame.sheetId;
  root.setAttribute("role", "grid"); root.setAttribute("aria-readonly", "true");
  root.setAttribute("aria-rowcount", "1048576"); root.setAttribute("aria-colcount", "16384");
  root.dir = rightToLeft ? "rtl" : "ltr";
  root.style.cssText = `position:relative;overflow:hidden;isolation:isolate;width:${frame.layout.widthPt}pt;height:${frame.layout.heightPt}pt`;
  for (const projection of frame.projections) {
    let row = rows.get(projection.row);
    if (!row) {
      row = document.createElement("div"); row.setAttribute("role", "row"); row.setAttribute("aria-rowindex", String(projection.row + 1));
      row.style.display = "contents"; rows.set(projection.row, row);
    }
    const clip = document.createElement("div"), body = document.createElement("div"), key = projection.row * 16_384 + projection.column;
    clip.dataset.row = String(projection.row); clip.dataset.column = String(projection.column); clip.dataset.pane = projection.pane;
    clip.dataset.officeView = "workbook-cell";
    if (accessible.has(key)) clip.setAttribute("aria-hidden", "true");
    else {
      accessible.add(key); clip.setAttribute("role", "gridcell"); clip.setAttribute("aria-colindex", String(projection.column + 1));
      if (projection.range.to.row !== projection.row) clip.setAttribute("aria-rowspan", String(projection.range.to.row - projection.row + 1));
      if (projection.range.to.column !== projection.column) clip.setAttribute("aria-colspan", String(projection.range.to.column - projection.column + 1));
    }
    const bounds = projection.clip, rect = projection.rect;
    const left = rightToLeft ? frame.layout.widthPt - bounds.x - bounds.width : bounds.x;
    const innerLeft = rightToLeft ? bounds.x + bounds.width - rect.x - rect.width : rect.x - bounds.x;
    clip.style.cssText = `position:absolute;overflow:hidden;left:${left}pt;top:${bounds.y}pt;width:${bounds.width}pt;height:${bounds.height}pt`;
    body.style.cssText = `position:absolute;box-sizing:border-box;left:${innerLeft}pt;top:${rect.y - bounds.y}pt;width:${rect.width}pt;height:${rect.height}pt`;
    paint(body, frame.getCell(projection.row, projection.column), projection);
    clip.append(body); row.append(clip);
  }
  for (const [, row] of [...rows].sort((a, b) => a[0] - b[0])) root.append(row);
  root.append(...borders);
  // No mutation of the current view until every projected cell was painted.
  container.replaceChildren(root);
  return root;
}
