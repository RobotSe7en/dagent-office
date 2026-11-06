import { DocumentError } from "@dagent-office/core";

/** Physical cell geometry. Ordinary borders anchor a rotated cell at its
 * bottom edge; diagonal-only borders do not change its rectangular outline.
 * The worksheet's column direction does not reverse the text rotation. */
export interface WorkbookCellShape {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly topOffsetPt: number;
}
export function workbookCellShape(widthPt: number, heightPt: number, rotationDegrees: number, bordered: boolean): WorkbookCellShape {
  if (![widthPt, heightPt, rotationDegrees].every(Number.isFinite) || widthPt <= 0 || heightPt <= 0 || Math.abs(rotationDegrees) > 90 || !Number.isInteger(rotationDegrees)
    || typeof bordered !== "boolean") throw new DocumentError("InvalidOperation", "A worksheet cell shape requires positive geometry and a whole-degree text rotation.");
  const topOffsetPt = !bordered || !rotationDegrees || Math.abs(rotationDegrees) === 90 ? 0 : -heightPt / Math.tan(rotationDegrees * Math.PI / 180);
  if (!Number.isFinite(topOffsetPt) || !Number.isFinite(widthPt + Math.abs(topOffsetPt))) throw new DocumentError("ResourceLimitExceeded", "A rotated worksheet cell exceeds finite geometry.");
  return Object.freeze({ widthPt, heightPt, topOffsetPt });
}
export function workbookCellShapeOffset(shape: WorkbookCellShape, yPt: number): number {
  if (!Number.isFinite(yPt)) throw new DocumentError("InvalidOperation", "A worksheet cell point must be finite.");
  const offset = shape.topOffsetPt * (1 - yPt / shape.heightPt);
  if (!Number.isFinite(offset)) throw new DocumentError("ResourceLimitExceeded", "A rotated worksheet point exceeds finite geometry.");
  return offset;
}
