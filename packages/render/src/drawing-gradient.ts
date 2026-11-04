import { DocumentError } from "@dagent-office/core";

export interface DrawingGradientBounds { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
export interface DrawingGradientLine { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number }

/** DrawingML lin/@scaled changes the physical gradient vector, not merely an
 * SVG bounding-box transform. Project all four fill corners onto that vector
 * so the stop interval covers the complete fill region in document units. */
export function drawingGradientLine(bounds: DrawingGradientBounds, angleDegrees: number, scaled: boolean): DrawingGradientLine {
  const { x, y, width, height } = bounds;
  if (![x, y, width, height, angleDegrees].every(Number.isFinite) || width < 0 || height < 0) throw new DocumentError("InvalidOperation", "A drawing gradient requires finite bounds and a finite angle.");
  const angle = (angleDegrees % 360) * Math.PI / 180, cosine = Math.cos(angle), sine = Math.sin(angle), scale = Math.max(width, height, 1);
  let nx = cosine * (scaled ? width / scale : 1), ny = sine * (scaled ? height / scale : 1);
  let length = Math.hypot(nx, ny);
  if (!length) { nx = cosine; ny = sine; length = 1; }
  nx /= length; ny /= length;
  const half = (Math.abs(nx) * width + Math.abs(ny) * height) / 2, cx = x + width / 2, cy = y + height / 2;
  const line = { x1: cx - nx * half, y1: cy - ny * half, x2: cx + nx * half, y2: cy + ny * half };
  if (!Object.values(line).every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "A drawing gradient exceeds finite coordinate geometry.");
  return Object.freeze(line);
}
