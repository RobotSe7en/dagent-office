import { DocumentError } from "@doffi/core";

export interface DrawingGradientBounds { readonly x: number; readonly y: number; readonly width: number; readonly height: number }
export interface DrawingGradientLine { readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number }
export interface DrawingGradientOrientation { readonly rotationDegrees: number; readonly flipH: boolean; readonly flipV: boolean }

/** DrawingML lin/@scaled changes the physical gradient vector, not merely an
 * SVG bounding-box transform. Project all four fill corners onto that vector
 * so the stop interval covers the complete fill region in document units.
 * fixedOrientation places the field in the page bounding rectangle and
 * cancels the owning compositor's rotation, retaining explicit reflection. */
export function drawingGradientLine(bounds: DrawingGradientBounds, angleDegrees: number, scaled: boolean, fixedOrientation?: DrawingGradientOrientation): DrawingGradientLine {
  const { x, y, width, height } = bounds;
  if (![x, y, width, height, angleDegrees].every(Number.isFinite) || width < 0 || height < 0) throw new DocumentError("InvalidOperation", "A drawing gradient requires finite bounds and a finite angle.");
  if (fixedOrientation) {
    const { rotationDegrees, flipH, flipV } = fixedOrientation;
    if (!Number.isFinite(rotationDegrees) || typeof flipH !== "boolean" || typeof flipV !== "boolean") throw new DocumentError("InvalidOperation", "A fixed gradient requires a finite object orientation.");
    if (rotationDegrees % 360 === 0) return drawingGradientLine(bounds, angleDegrees, scaled);
    const degrees = rotationDegrees % 360, radians = degrees * Math.PI / 180;
    const cosine = degrees % 90 === 0 ? (degrees % 180 === 0 ? 1 : 0) : Math.abs(Math.cos(radians));
    const sine = degrees % 90 === 0 ? (degrees % 180 === 0 ? 0 : 1) : Math.abs(Math.sin(radians));
    const w = width * cosine + height * sine, h = width * sine + height * cosine, cx = x + width / 2, cy = y + height / 2;
    if (![w, h, cx, cy].every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "A fixed gradient exceeds its rotated coordinate bounds.");
    const line = drawingGradientLine({ x: cx - w / 2, y: cy - h / 2, width: w, height: h }, angleDegrees, scaled);
    // The compositor applies R*F. Counter-rotation conjugated by F leaves
    // the fill in page axes while preserving the source's explicit reflection.
    const counter = radians * (flipH !== flipV ? 1 : -1);
    const c = degrees % 90 === 0 ? (degrees % 180 === 0 ? Math.sign(Math.cos(counter)) : 0) : Math.cos(counter);
    const s = degrees % 90 === 0 ? (degrees % 180 === 0 ? 0 : Math.sign(Math.sin(counter))) : Math.sin(counter);
    const project = (px: number, py: number) => ({ x: cx + (px - cx) * c - (py - cy) * s, y: cy + (px - cx) * s + (py - cy) * c });
    const first = project(line.x1, line.y1), last = project(line.x2, line.y2), result = { x1: first.x, y1: first.y, x2: last.x, y2: last.y };
    if (!Object.values(result).every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "A fixed gradient exceeds finite coordinate geometry.");
    return Object.freeze(result);
  }
  const degrees = angleDegrees % 360, angle = degrees * Math.PI / 180;
  const cosine = degrees % 90 === 0 ? degrees % 180 === 0 ? Math.sign(Math.cos(angle)) : 0 : Math.cos(angle);
  const sine = degrees % 90 === 0 ? degrees % 180 === 0 ? 0 : Math.sign(Math.sin(angle)) : Math.sin(angle), scale = Math.max(width, height, 1);
  let nx = cosine * (scaled ? width / scale : 1), ny = sine * (scaled ? height / scale : 1);
  let length = Math.hypot(nx, ny);
  if (!length) { nx = cosine; ny = sine; length = 1; }
  nx /= length; ny /= length;
  const half = (Math.abs(nx) * width + Math.abs(ny) * height) / 2, cx = x + width / 2, cy = y + height / 2;
  const line = { x1: cx - nx * half, y1: cy - ny * half, x2: cx + nx * half, y2: cy + ny * half };
  if (!Object.values(line).every(Number.isFinite)) throw new DocumentError("ResourceLimitExceeded", "A drawing gradient exceeds finite coordinate geometry.");
  return Object.freeze(line);
}
