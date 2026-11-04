import { DocumentError } from "@dagent-office/core";
import type { PresentationPresetGeometry } from "@dagent-office/core";
import { addArrows, strokeAttributes, svgFill } from "./presentation/appearance.js";
import { svgElement, svgPath } from "./presentation/svg.js";
import { markViewNode } from "./view-nodes.js";

export interface DrawingGeometryViewInput {
  readonly preset: string; readonly geometry: PresentationPresetGeometry;
  readonly widthPt: number; readonly heightPt: number; readonly id: string;
  readonly fillFallback?: string; readonly strokeFallback?: string; readonly maxNodes?: number;
}
/** Shared SVG projection of first-party preset paths. Layout, source ownership,
 * text rectangles and object transforms remain with each format's compositor. */
export function createDrawingGeometryView(document: Document, style: CSSStyleDeclaration, input: DrawingGeometryViewInput): SVGSVGElement {
  const { widthPt, heightPt, id } = input, maxNodes = input.maxNodes ?? 50_000;
  if (![widthPt, heightPt].every(value => Number.isFinite(value) && value >= 0) || !Number.isSafeInteger(maxNodes) || maxNodes < 0 || !/^[a-z][a-z0-9-]*$/i.test(id)) throw new DocumentError("InvalidOperation", "Drawing geometry requires finite extents, a node budget and an isolated SVG identity.");
  let count = 0;
  const account = (amount: number): void => { count += amount; if (count > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing geometry exceeds its node budget."); };
  account(1);
  const svg = markViewNode(svgElement(document, "svg", { width: "100%", height: "100%", viewBox: `0 0 ${Math.max(widthPt, .001)} ${Math.max(heightPt, .001)}`, "aria-hidden": "true", "data-sdk-preset": input.preset }));
  svg.style.cssText = `position:absolute;left:0;top:0;width:${Math.max(widthPt, .001)}pt;height:${Math.max(heightPt, .001)}pt;overflow:visible;pointer-events:none`;
  const fill = svgFill(document, style, svg, id + "-fill", "--office-fill", input.fillFallback ?? "none", account);
  for (const [index, path] of input.geometry.paths.entries()) {
    account(1);
    const drawing = svgElement(document, "path", { d: svgPath(path), fill: path.fill ? fill : "none", ...strokeAttributes(document, style, "--office-line", input.strokeFallback ?? "none") });
    if (!path.stroke) drawing.setAttribute("stroke", "none");
    // Multi-path shapes must not publish duplicate marker IDs in one SVG.
    addArrows(document, svg, drawing, style, `${id}-path-${index}`, account); svg.append(drawing);
  }
  return svg;
}
