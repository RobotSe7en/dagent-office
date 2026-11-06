import type { ChartAppearance, DeepReadonly } from "@doffi/core";
import { DocumentError } from "@doffi/core";
import { svgElement } from "./presentation/svg.js";

type Marker = Exclude<NonNullable<DeepReadonly<ChartAppearance>["seriesStyles"]>[number]["marker"], "auto">;
/** ECMA-376 §21.2.3.27: dash/dot are filled rectangles; open symbols have
 * a square fill beneath their internal strokes. All coordinates are points. */
export function appendChartMarker(svg: SVGElement, marker: Marker, size: number, x: number, y: number,
  fill: string, stroke: Readonly<Record<string, string | number>>, create: typeof svgElement): SVGGElement | undefined {
  if (!marker || marker === "none") return;
  if (!Number.isFinite(size) || size <= 0 || size > 1_000_000) throw new DocumentError("ResourceLimitExceeded", "A chart marker exceeds its finite coordinate budget.");
  const document = svg.ownerDocument, r = size / 2;
  const group = create(document, "g", { "data-sdk-chart-marker": marker }); svg.append(group);
  if (marker === "circle") group.append(create(document, "circle", { cx: x, cy: y, r, fill, ...stroke }));
  else if (marker === "square" || marker === "dash" || marker === "dot") {
    const width = marker === "dot" ? size / 2 : size, height = marker === "square" ? size : size / 5;
    group.append(create(document, "rect", { x: x - width / 2, y: y - height / 2, width, height, fill, ...stroke }));
  } else if (marker === "triangle" || marker === "diamond") {
    const d = marker === "triangle" ? `M${x} ${y - r}L${x + r} ${y + r}L${x - r} ${y + r}Z` : `M${x} ${y - r}L${x + r} ${y}L${x} ${y + r}L${x - r} ${y}Z`;
    group.append(create(document, "path", { d, fill, ...stroke }));
  } else {
    group.append(create(document, "rect", { x: x - r, y: y - r, width: size, height: size, fill, stroke: "none" }));
    const plus = `M${x - r} ${y}L${x + r} ${y}M${x} ${y - r}L${x} ${y + r}`;
    const cross = `M${x - r} ${y - r}L${x + r} ${y + r}M${x - r} ${y + r}L${x + r} ${y - r}`;
    group.append(create(document, "path", { d: marker === "star" ? plus + cross : marker === "x" ? cross : plus, fill: "none", ...stroke }));
  }
  return group;
}
