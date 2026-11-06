import { DocumentError } from "@doffi/core";
import { drawingGradientLine } from "../drawing-gradient.js";
import type { DrawingGradientLine } from "../drawing-gradient.js";
import { svgElement } from "../presentation/svg.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookCellFill, WorkbookGradientStop } from "./cell-fill.js";
import { workbookFillPattern } from "./fill-pattern.js";

export interface WorkbookFillClip { readonly xPt: number; readonly yPt: number; readonly widthPt: number; readonly heightPt: number }
type Point = readonly [number, number];
function clippedPolygon(points: readonly Point[], clip: WorkbookFillClip): readonly Point[] {
  let result = [...points];
  for (const [axis, bound, minimum] of [[0, clip.xPt, true], [0, clip.xPt + clip.widthPt, false], [1, clip.yPt, true], [1, clip.yPt + clip.heightPt, false]] as const) {
    const input = result; result = [];
    for (let i = 0; i < input.length; i++) {
      const a = input[i]!, b = input[(i + 1) % input.length]!, insideA = minimum ? a[axis] >= bound : a[axis] <= bound, insideB = minimum ? b[axis] >= bound : b[axis] <= bound;
      if (insideA) result.push(a);
      if (insideA !== insideB) {
        const fraction = (bound - a[axis]) / (b[axis] - a[axis]);
        result.push(axis === 0 ? [bound, a[1] + (b[1] - a[1]) * fraction] : [a[0] + (b[0] - a[0]) * fraction, bound]);
      }
    }
  }
  return result.map(([x, y]) => [x - clip.xPt, y - clip.yPt]);
}

/** Paint a visible fragment in document points. Pattern phase and gradient
 * geometry belong to the full cell; scrolling/freeze panes never restart them.
 * Large logical cells do not allocate a cell-sized bitmap or DOM surface. */
export function createWorkbookFillView(document: Document, paint: WorkbookCellFill, widthPt: number, heightPt: number, id: string,
  clip: WorkbookFillClip = { xPt: 0, yPt: 0, widthPt, heightPt }, maxNodes = 50_000): { readonly element: SVGSVGElement; readonly nodes: number } {
  if (![widthPt, heightPt, clip.xPt, clip.yPt, clip.widthPt, clip.heightPt].every(Number.isFinite) || widthPt <= 0 || heightPt <= 0 || clip.xPt < 0 || clip.yPt < 0
    || clip.widthPt <= 0 || clip.heightPt <= 0 || clip.xPt + clip.widthPt > widthPt || clip.yPt + clip.heightPt > heightPt || !/^[A-Za-z_][A-Za-z0-9_.-]*$/.test(id)
    || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "A worksheet fill needs finite cell geometry, a contained visible fragment and valid budgets.");
  if (clip.widthPt > 1_000_000 || clip.heightPt > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "A worksheet fill fragment exceeds the bounded view extent.");
  let nodes = 0;
  const make: typeof svgElement = (doc, tag, attributes) => {
    if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet fill geometry exceeds its node budget.");
    return markViewNode(svgElement(doc, tag, attributes));
  };
  const element = make(document, "svg", { viewBox: `0 0 ${clip.widthPt} ${clip.heightPt}`, "aria-hidden": "true", "data-office-view": "workbook-cell-fill" }) as SVGSVGElement;
  element.style.cssText = `position:absolute;left:0;top:0;width:${clip.widthPt}pt;height:${clip.heightPt}pt;pointer-events:none;overflow:hidden`;
  const rect = (fill: string): void => { element.append(make(document, "rect", { x: 0, y: 0, width: clip.widthPt, height: clip.heightPt, fill })); };
  try {
    if (paint.kind === "solid") rect(paint.color);
    else if (paint.kind === "pattern") {
      const rows = workbookFillPattern(paint.pattern), width = rows[0]!.length * .75, height = rows.length * .75;
      const defs = make(document, "defs", {}), pattern = make(document, "pattern", { id, patternUnits: "userSpaceOnUse", width, height, x: -(clip.xPt % width), y: -(clip.yPt % height) });
      pattern.append(make(document, "rect", { x: 0, y: 0, width, height, fill: paint.background }));
      let d = "";
      for (const [y, row] of rows.entries()) for (let x = 0; x < row.length; x++) if (row[x] === "1") d += `M${x * .75} ${y * .75}h.75v.75h-.75Z`;
      pattern.append(make(document, "path", { d, fill: paint.foreground, "shape-rendering": "crispEdges" })); defs.append(pattern); element.append(defs); rect(`url(#${id})`);
    } else if (paint.stops.length) {
      if (paint.stops.length === 1) rect(paint.stops[0]!.color);
      else {
        const defs = make(document, "defs", {}); element.append(defs);
        const gradient = (name: string, line: DrawingGradientLine, stops: readonly WorkbookGradientStop[]): string => {
          const node = make(document, "linearGradient", { id: name, gradientUnits: "userSpaceOnUse", x1: line.x1 - clip.xPt, y1: line.y1 - clip.yPt, x2: line.x2 - clip.xPt, y2: line.y2 - clip.yPt, "color-interpolation": "sRGB" });
          for (const stop of stops) node.append(make(document, "stop", { offset: stop.position, "stop-color": stop.color }));
          defs.append(node); return `url(#${name})`;
        };
        if (paint.kind === "linear") rect(gradient(id, drawingGradientLine({ x: 0, y: 0, width: widthPt, height: heightPt }, paint.degree, false), paint.stops));
        else {
          const left = Math.min(paint.left, paint.right) * widthPt, right = Math.max(paint.left, paint.right) * widthPt;
          const top = Math.min(paint.top, paint.bottom) * heightPt, bottom = Math.max(paint.top, paint.bottom) * heightPt;
          // Four linear trapezoids produce rectangular equal-color contours;
          // a radial gradient would turn the specified rectangle into an ellipse.
          const pieces: { points: readonly Point[]; line: DrawingGradientLine; size: number }[] = [
            { points: [[0, 0], [left, top], [left, bottom], [0, heightPt]], line: { x1: left, y1: 0, x2: 0, y2: 0 }, size: left },
            { points: [[right, top], [widthPt, 0], [widthPt, heightPt], [right, bottom]], line: { x1: right, y1: 0, x2: widthPt, y2: 0 }, size: widthPt - right },
            { points: [[0, 0], [widthPt, 0], [right, top], [left, top]], line: { x1: 0, y1: top, x2: 0, y2: 0 }, size: top },
            { points: [[left, bottom], [right, bottom], [widthPt, heightPt], [0, heightPt]], line: { x1: 0, y1: bottom, x2: 0, y2: heightPt }, size: heightPt - bottom },
          ];
          rect(paint.stops.at(-1)!.color);
          for (const [index, piece] of pieces.entries()) {
            const points = piece.size > 0 ? clippedPolygon(piece.points, clip) : [];
            // Adjacent pieces meet at equal colors. Rasterizing their shared
            // boundary without alpha prevents the base color bleeding through.
            if (points.length >= 3) element.append(make(document, "polygon", { points: points.map(point => point.join(",")).join(" "), fill: gradient(id + "-" + index, piece.line, paint.stops), "shape-rendering": "crispEdges" }));
          }
          const x = Math.max(left, clip.xPt), y = Math.max(top, clip.yPt), w = Math.min(right, clip.xPt + clip.widthPt) - x, h = Math.min(bottom, clip.yPt + clip.heightPt) - y;
          if (w > 0 && h > 0) element.append(make(document, "rect", { x: x - clip.xPt, y: y - clip.yPt, width: w, height: h, fill: paint.stops[0]!.color, "shape-rendering": "crispEdges" }));
        }
      }
    }
    return Object.freeze({ element, nodes });
  } catch (error) { element.replaceChildren(); throw error; }
}
