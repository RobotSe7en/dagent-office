import { DocumentError, evaluatePresentationPreset } from "@doffi/core";
import type { DeepReadonly, PresentationPresetGeometry, WorkbookObject } from "@doffi/core";
import { createDrawingGeometryView } from "../drawing-geometry.js";
import { workbookDrawingOrientation } from "./drawing-view.js";

export function createWorkbookObjectGeometry(source: HTMLElement, object: DeepReadonly<WorkbookObject>, widthPt: number, heightPt: number, id: string, maxNodes = 50_000): { readonly element: SVGSVGElement; readonly geometry: PresentationPresetGeometry } {
  const document = source.ownerDocument;
  if (!source.isConnected || !document.defaultView) throw new DocumentError("InvalidOperation", "Drawing appearance requires a connected source projection.");
  if (source.dataset.nodeId !== object.nodeId || source.dataset.officeKind !== object.kind || object.kind === "chart") throw new DocumentError("InvalidDocument", "Drawing geometry requires its own non-chart source object.");
  const preset = object.kind === "geometry" ? object.preset : "rect", adjustments: Record<string, number> = {};
  for (const attribute of source.attributes) if (attribute.name.startsWith("data-office-adjust-")) adjustments[attribute.name.slice(19)] = Number(attribute.value);
  const geometry = evaluatePresentationPreset(preset, widthPt, heightPt, adjustments), style = document.defaultView.getComputedStyle(source);
  const background = style.backgroundColor, hasBackground = background !== "transparent" && !/rgba\([^)]*,\s*0\)$/.test(background);
  const element = createDrawingGeometryView(document, style, { preset, geometry, widthPt, heightPt, id, maxNodes, orientation: workbookDrawingOrientation(source),
    fillFallback: preset === "line" ? "none" : hasBackground ? background : object.kind === "geometry" ? "#4472c4" : "none",
    strokeFallback: preset === "line" ? "#000000" : "none" });
  return Object.freeze({ element, geometry });
}
