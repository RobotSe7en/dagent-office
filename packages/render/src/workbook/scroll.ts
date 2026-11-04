import { DocumentError } from "@dagent-office/core";

/** Keep a bounded native scrollbar while the worksheet retains its full point
 * extent. Wheel/keyboard movement uses logical points directly; only scrollbar
 * thumb movement is scaled. No document dimension is rewritten to screen pixels. */
export class WorkbookScrollAxis {
  readonly maximumPt: number;
  readonly trackMaximumPx: number;
  readonly trackSizePx: number;
  constructor(extentPt: number, viewportPt: number, trackLimitPx = 1_000_000) {
    if (![extentPt, viewportPt, trackLimitPx].every(Number.isFinite) || extentPt < 0 || viewportPt < 0 || trackLimitPx <= 0)
      throw new DocumentError("InvalidOperation", "Invalid logical workbook scroll dimensions.");
    this.maximumPt = Math.max(0, extentPt - viewportPt);
    this.trackMaximumPx = Math.min(this.maximumPt / 0.75, trackLimitPx);
    this.trackSizePx = viewportPt / 0.75 + this.trackMaximumPx;
    if (!Number.isFinite(this.trackSizePx)) throw new DocumentError("InvalidOperation", "Workbook scrollbar dimensions exceed their numeric range.");
  }
  clamp(point: number): number {
    if (!Number.isFinite(point)) throw new DocumentError("InvalidOperation", "Workbook scroll positions must be finite.");
    return Math.max(0, Math.min(point, this.maximumPt));
  }
  toTrack(point: number): number { return this.maximumPt ? this.clamp(point) / this.maximumPt * this.trackMaximumPx : this.clamp(point); }
  fromTrack(pixel: number): number {
    if (!Number.isFinite(pixel)) throw new DocumentError("InvalidOperation", "Workbook scrollbar positions must be finite.");
    return this.trackMaximumPx ? Math.max(0, Math.min(pixel, this.trackMaximumPx)) / this.trackMaximumPx * this.maximumPt : 0;
  }
}
