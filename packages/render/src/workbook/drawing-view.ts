import { DocumentError } from "@doffi/core";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookDrawingProjection } from "./drawing.js";

export interface WorkbookDrawingOrientation { readonly rotationDegrees: number; readonly flipH: boolean; readonly flipV: boolean; readonly hidden: boolean }
export function workbookDrawingOrientation(source: HTMLElement): WorkbookDrawingOrientation {
  const rotation = source.getAttribute("data-office-rotation"), rotationDegrees = rotation === null ? 0 : Number(rotation);
  if (rotation !== null && !/^-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(rotation) || !Number.isFinite(rotationDegrees)) throw new DocumentError("InvalidDocument", "A drawing rotation must be a finite source number.");
  const flag = (name: string): boolean => {
    const value = source.getAttribute("data-office-" + name);
    if (value !== null && value !== "false" && value !== "true") throw new DocumentError("InvalidDocument", "A drawing orientation flag must be a source boolean.");
    return value === "true";
  };
  return Object.freeze({ rotationDegrees, flipH: flag("flip-h"), flipV: flag("flip-v"), hidden: flag("hidden") });
}

/** A pane clips a fully transformed object. RTL changes its physical anchor
 * position; explicit source flips alone mirror its contents. A separate fresh
 * content subtree is required for each pane, with shared leased image bytes. */
export function createWorkbookDrawingFragment(document: Document, nodeId: string, projection: WorkbookDrawingProjection, viewportWidthPt: number, orientation: WorkbookDrawingOrientation, rightToLeft: boolean, content: HTMLElement): HTMLElement {
  const { rect, clip } = projection;
  if (![rect.x, rect.y, rect.width, rect.height, clip.x, clip.y, clip.width, clip.height, viewportWidthPt, orientation.rotationDegrees].every(Number.isFinite)
    || Math.min(rect.width, rect.height) < 0 || Math.min(clip.width, clip.height, viewportWidthPt) <= 0 || content.ownerDocument !== document || content.parentNode || content.isConnected)
    throw new DocumentError("InvalidOperation", "A drawing fragment requires finite geometry and fresh detached content.");
  const root = markViewNode(document.createElement("div")), body = document.createElement("div");
  root.dataset.officeView = "workbook-drawing"; root.dataset.sourceNodeId = nodeId; root.dataset.pane = projection.pane;
  const left = rightToLeft ? viewportWidthPt - clip.x - clip.width : clip.x;
  const innerLeft = rightToLeft ? clip.x + clip.width - rect.x - rect.width : rect.x - clip.x;
  root.style.cssText = `position:absolute;overflow:hidden;left:${left}pt;top:${clip.y}pt;width:${clip.width}pt;height:${clip.height}pt;direction:ltr`;
  body.dataset.officeView = "workbook-drawing-body";
  body.style.cssText = `position:absolute;box-sizing:border-box;left:${innerLeft}pt;top:${rect.y - clip.y}pt;width:${rect.width}pt;height:${rect.height}pt;transform-origin:50% 50%;transform:rotate(${orientation.rotationDegrees % 360}deg) scale(${orientation.flipH ? -1 : 1},${orientation.flipV ? -1 : 1})`;
  if (orientation.hidden) root.style.visibility = "hidden";
  body.append(content); root.append(body); return root;
}
