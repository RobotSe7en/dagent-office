import { DocumentError } from "@dagent-office/core";
import { property, resolvedColor } from "../presentation/appearance.js";
import type { WorkbookDrawingOrientation } from "./drawing-view.js";

export interface WorkbookDrawingShadow {
  readonly filter: string;
  readonly localOffsetPt: { readonly x: number; readonly y: number };
  readonly outsetPt: number;
  readonly rasterPixels: number;
  /** Native blur radius is retained in source; CSS uses a Gaussian approximation. */
  readonly approximateBlur: boolean;
}
export interface WorkbookShadowLimits { readonly maxRasterPixels?: number; readonly pixelsPerPoint?: number }

/** Resolve an outer shadow for the untransformed object-content group. The
 * fragment applies rotation/reflection afterwards. Page-oriented effects
 * therefore need the inverse object transform, not merely a negated angle.
 * Scale/skew shadows are rejected by the native appearance reader. Callers
 * supply the full painted bounds (including outline/text outsets), add the
 * returned outset to pane culling, and accumulate rasterPixels across their
 * candidate. Per-object admission does not reset the whole-window budget. */
export function workbookDrawingShadow(source: HTMLElement, orientation: WorkbookDrawingOrientation, widthPt: number, heightPt: number, limits: WorkbookShadowLimits = {}): WorkbookDrawingShadow {
  const document = source.ownerDocument, view = document.defaultView, maxPixels = limits.maxRasterPixels ?? 16_777_216, pixelsPerPoint = limits.pixelsPerPoint ?? (view?.devicePixelRatio ?? 1) / .75;
  if (!source.isConnected || !view || ![widthPt, heightPt, orientation.rotationDegrees].every(Number.isFinite) || Math.min(widthPt, heightPt) < 0
    || !Number.isSafeInteger(maxPixels) || maxPixels < 1 || !Number.isFinite(pixelsPerPoint) || pixelsPerPoint <= 0) throw new DocumentError("InvalidOperation", "Drawing shadows require connected source styles, finite geometry and positive raster limits.");
  const style = view.getComputedStyle(source), kind = property(style, "--office-shadow-kind", "none");
  if (kind === "none") return Object.freeze({ filter: "none", localOffsetPt: Object.freeze({ x: 0, y: 0 }), outsetPt: 0, rasterPixels: 0, approximateBlur: false });
  if (kind !== "outer") throw new DocumentError("UnsupportedFeature", "This drawing shadow kind is not supported.");
  const number = (name: string, length = false): number => {
    const raw = property(style, "--office-shadow-" + name, length ? "0pt" : "0");
    const match = /^(-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?)(pt)?$/i.exec(raw), value = match ? Number(match[1]) : NaN;
    if (!match || !Number.isFinite(value) || length && (value < 0 || !match[2] && value !== 0) || !length && match[2]) throw new DocumentError("InvalidDocument", "A drawing shadow has an invalid source number or point length.");
    return value;
  };
  const blur = number("blur", true), distance = number("distance", true), direction = number("direction") % 360;
  const rotate = property(style, "--office-shadow-rotate-with-shape", "true"), align = property(style, "--office-shadow-align", "b");
  if (!["true", "false"].includes(rotate) || !["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"].includes(align)) throw new DocumentError("InvalidDocument", "A drawing shadow orientation or alignment is invalid.");
  const angle = direction * Math.PI / 180, x = distance * Math.cos(angle), y = distance * Math.sin(angle), rotation = orientation.rotationDegrees % 360 * Math.PI / 180;
  const local = rotate === "true" ? { x, y } : { x: (orientation.flipH ? -1 : 1) * (Math.cos(rotation) * x + Math.sin(rotation) * y), y: (orientation.flipV ? -1 : 1) * (-Math.sin(rotation) * x + Math.cos(rotation) * y) };
  const sigma = blur / 2, outset = distance + 3 * sigma;
  const rasterWidth = Math.ceil((widthPt + 2 * outset) * pixelsPerPoint), rasterHeight = Math.ceil((heightPt + 2 * outset) * pixelsPerPoint), pixels = Math.max(1, rasterWidth) * Math.max(1, rasterHeight);
  if (![outset, rasterWidth, rasterHeight, pixels, local.x, local.y].every(Number.isFinite) || !Number.isSafeInteger(pixels) || pixels > maxPixels) throw new DocumentError("ResourceLimitExceeded", "A drawing shadow exceeds its raster budget.");
  const color = resolvedColor(document, style, "--office-shadow");
  return Object.freeze({ filter: `drop-shadow(${local.x}pt ${local.y}pt ${sigma}pt ${color})`, localOffsetPt: Object.freeze(local), outsetPt: outset, rasterPixels: pixels, approximateBlur: blur > 0 });
}
