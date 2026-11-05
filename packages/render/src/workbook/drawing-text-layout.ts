import { DocumentError, evaluatePresentationPreset } from "@dagent-office/core";
import { measuredFontLineHeight } from "../font-line-height.js";
import type { DeepReadonly, DrawingTextProperties, PresentationPresetGeometry, WorkbookChart, WorkbookObject } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { markViewNode } from "../view-nodes.js";
import { property } from "../presentation/appearance.js";
import { decorateDrawingLists } from "../presentation/text.js";
import { workbookDrawingOrientation } from "./drawing-view.js";
import { workbookChartTextSource } from "./drawing-text.js";
import type { WorkbookChartTextRole, WorkbookDrawingTextView } from "./drawing-text.js";

export interface WorkbookDrawingTextLayoutInput {
  readonly source: HTMLElement; readonly object: DeepReadonly<WorkbookObject>; readonly text: WorkbookDrawingTextView;
  readonly widthPt: number; readonly heightPt: number;
}
export interface WorkbookDrawingTextLayout {
  readonly element: HTMLElement; readonly text: WorkbookDrawingTextView; readonly widthPt: number; readonly heightPt: number;
  readonly geometry: PresentationPresetGeometry; readonly scale: number; readonly overflow: boolean;
  readonly nodes: number;
}
export interface WorkbookDrawingTextLimits { readonly maxNodes?: number; readonly maxPasses?: number; readonly maxExtentPt?: number }
export interface WorkbookDrawingTextBatch {
  readonly objects: readonly WorkbookDrawingTextLayout[]; readonly fonts: readonly FontResolution[]; readonly usedFallbackLineMetrics: boolean;
}
export interface WorkbookChartTextLayoutInput {
  readonly source: HTMLElement; readonly chart: DeepReadonly<WorkbookChart>; readonly role: WorkbookChartTextRole; readonly text: WorkbookDrawingTextView;
  readonly widthPt: number; readonly heightPt: number;
}
interface TextBodyLayoutInput {
  readonly source: HTMLElement; readonly text: WorkbookDrawingTextView; readonly widthPt: number; readonly heightPt: number;
  readonly properties: DeepReadonly<DrawingTextProperties> | undefined;
  readonly geometry: (width: number, height: number) => PresentationPresetGeometry;
  readonly flipH: boolean; readonly flipV: boolean;
}

/** Resolve local fonts and measure complete, unrotated object rectangles before
 * their pane projections are created. Text counters its object's reflections;
 * the outer object still owns rotation, flips, fill/effects and clipping. */
export function layoutWorkbookDrawingTexts(scope: HTMLElement, inputs: readonly WorkbookDrawingTextLayoutInput[], limits: WorkbookDrawingTextLimits = {}): WorkbookDrawingTextBatch {
  return layoutTextBodies(scope, inputs, limits, input => {
    const { object, source } = input;
    if (object.kind !== "text" && object.kind !== "geometry" || source.dataset.nodeId !== object.nodeId || source.dataset.officeKind !== object.kind) throw new DocumentError("InvalidOperation", "Drawing layout requires its own shape or text-box source.");
    const orientation = workbookDrawingOrientation(source), adjustments: Record<string, number> = {};
    for (const attribute of source.attributes) if (attribute.name.startsWith("data-office-adjust-")) adjustments[attribute.name.slice(19)] = Number(attribute.value);
    const preset = object.kind === "geometry" ? object.preset : "rect";
    return { source, text: input.text, widthPt: input.widthPt, heightPt: input.heightPt, properties: object.text,
      geometry: (width: number, height: number) => evaluatePresentationPreset(preset, width, height, adjustments), flipH: orientation.flipH, flipV: orientation.flipV };
  });
}

/** Measure each rich title in a compositor-supplied local rectangle. The chart
 * compositor owns title placement, axis-title rotation and surrounding graphics;
 * the title keeps its real source identity and DrawingML text-body properties. */
export function layoutWorkbookChartTexts(scope: HTMLElement, inputs: readonly WorkbookChartTextLayoutInput[], limits: WorkbookDrawingTextLimits = {}): WorkbookDrawingTextBatch {
  return layoutTextBodies(scope, inputs, limits, input => {
    const { source, properties } = workbookChartTextSource(input.source, input.chart, input.role);
    return { source, properties, text: input.text, widthPt: input.widthPt, heightPt: input.heightPt,
      geometry: (width: number, height: number) => evaluatePresentationPreset("rect", width, height), flipH: false, flipV: false };
  });
}

function layoutTextBodies<T>(scope: HTMLElement, inputs: readonly T[], limits: WorkbookDrawingTextLimits, adapt: (input: T) => TextBodyLayoutInput): WorkbookDrawingTextBatch {
  const document = scope.ownerDocument, view = document.defaultView, maxNodes = limits.maxNodes ?? 50_000, maxPasses = limits.maxPasses ?? 8, maxExtent = limits.maxExtentPt ?? 1_000_000;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isSafeInteger(maxNodes) || maxNodes < 1 || !Number.isSafeInteger(maxPasses) || maxPasses < 1 || maxPasses > 128 || !Number.isFinite(maxExtent) || maxExtent <= 0) throw new DocumentError("InvalidOperation", "Drawing text layout requires a connected surface, settled fonts and finite layout budgets.");
  let nodes = 1;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Drawing text layout exceeds its node budget."); };
  const seen = new Set<HTMLElement>(), bodies: TextBodyLayoutInput[] = [];
  for (const original of inputs) {
    const input = adapt(original);
    if (input.source.ownerDocument !== document || !input.source.isConnected || input.text.element.ownerDocument !== document || input.text.element.parentNode || seen.has(input.text.element)
      || input.text.element.dataset.sourceNodeId !== input.source.dataset.nodeId
      || ![input.widthPt, input.heightPt].every(n => Number.isFinite(n) && n >= 0)) throw new DocumentError("InvalidOperation", "Drawing layout requires its own connected source, distinct detached text and finite object extents.");
    if (Math.max(input.widthPt, input.heightPt) > maxExtent) throw new DocumentError("ResourceLimitExceeded", "Drawing text layout exceeds its extent budget.");
    seen.add(input.text.element); account(input.text.nodes + 1); bodies.push(input);
  }
  const root = markViewNode(document.createElement("div")); root.dataset.officeView = "workbook-drawing-text-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden;direction:ltr";
  let complete = false;
  try {
    const prepared = bodies.map(input => {
      const computed = view.getComputedStyle(input.source), text = input.text.element, element = markViewNode(document.createElement("div"));
      const anchor = property(computed, "--office-text-anchor", "t"), centered = property(computed, "--office-text-anchor-center", "false");
      if (!["t", "ctr", "b", "just", "dist"].includes(anchor) || !["false", "true"].includes(centered)) throw new DocumentError("InvalidDocument", "A drawing text anchor is invalid.");
      if (centered === "true") throw new DocumentError("UnsupportedFeature", "Centered drawing text-block anchoring requires measured line geometry.");
      element.style.cssText = "position:relative;direction:ltr";
      Object.assign(text.style, { position: "absolute", display: "flex", flexDirection: "column", boxSizing: "border-box", margin: "0", overflow: "visible", justifyContent: "flex-start",
        whiteSpace: input.properties?.wrap === false ? "pre" : "pre-wrap", overflowWrap: input.properties?.wrap === false ? "normal" : "anywhere", transformOrigin: "50% 50%", transform: `scale(${input.flipH ? -1 : 1},${input.flipV ? -1 : 1})` });
      for (const side of ["left", "right", "top", "bottom"]) text.style.setProperty("padding-" + side, computed.getPropertyValue("padding-" + side));
      element.append(text); root.append(element);
      const state = { input, element, text, anchor, nodes: input.text.nodes + 1, width: input.widthPt, height: input.heightPt, geometry: input.geometry(input.widthPt, input.heightPt),
        resize(width: number, height: number) {
          if (![width, height].every(n => Number.isFinite(n) && n >= 0 && n <= maxExtent)) throw new DocumentError("ResourceLimitExceeded", "Drawing text auto-fit exceeds its extent budget.");
          this.width = width; this.height = height; this.geometry = input.geometry(width, height);
          element.style.width = width + "pt"; element.style.height = height + "pt";
          const rect = this.geometry.textRect;
          Object.assign(text.style, rect ? { display: "flex", left: rect.x + "pt", top: rect.y + "pt", width: rect.width + "pt", height: rect.height + "pt" } : { display: "none" });
        } };
      state.resize(state.width, state.height); return state;
    });
    scope.append(root);
    for (const item of prepared) decorateDrawingLists(item.text, amount => { account(amount); item.nodes += amount; });
    const fonts = resolvePreviewFontsInScope(document, root, "[data-office-text-style],[data-office-script],.sdk-ppt-marker");
    const normalLines = new Map<string, number>(), lineContext = document.createElement("canvas").getContext("2d");
    let usedFallbackLineMetrics = false;
    const lineHeight = (computed: CSSStyleDeclaration): number => {
      const line = parseFloat(computed.lineHeight); if (Number.isFinite(line)) return line;
      if (parseFloat(computed.fontSize) === 0) return 0;
      const properties = ["font-family", "font-size", "font-weight", "font-style", "font-stretch", "font-variant", "letter-spacing"], key = JSON.stringify(properties.map(name => computed.getPropertyValue(name)));
      const cached = normalLines.get(key); if (cached !== undefined) return cached;
      // Measure scalable font bounds at a fixed large size: browser pixel
      // rounding at the actual 14–18pt size otherwise adds a full pixel to
      // every line in some engines. No large canvas bitmap is allocated.
      const height = measuredFontLineHeight(lineContext, computed);
      if (height !== undefined) { normalLines.set(key, height); return height; }
      // Legacy engines do not expose font bounds. Keep a measured DOM fallback
      // and report its use; do not silently claim native normal-leading parity.
      usedFallbackLineMetrics = true;
      account(2); const probe = document.createElement("span");
      probe.style.cssText = "position:absolute;display:block;width:max-content;white-space:pre;line-height:normal;margin:0;padding:0;border:0";
      for (const name of properties) probe.style.setProperty(name, computed.getPropertyValue(name));
      probe.textContent = "M"; root.append(probe);
      try {
        const height = probe.getBoundingClientRect().height;
        if (!Number.isFinite(height) || height < 0 || height === 0 && parseFloat(computed.fontSize) > 0) throw new DocumentError("UnsupportedFeature", "This browser cannot measure the resolved drawing font's normal leading.");
        normalLines.set(key, height); return height;
      } finally { probe.remove(); nodes -= 2; }
    };
    const objects = prepared.map(item => {
      const { text } = item, properties = item.input.properties, styles = [text, ...text.querySelectorAll<HTMLElement>("[data-office-text-style],[data-office-script],.sdk-ppt-marker")].map(node => {
        const computed = view.getComputedStyle(node); return { node, size: parseFloat(computed.fontSize), line: lineHeight(computed) };
      });
      const apply = (scale: number): void => { for (const { node, size, line } of styles) {
        // DrawingML's paragraph default is an inheritance/insertion style,
        // not an extra large glyph on every line. Actual run/empty-line boxes
        // carry explicit sizes and leading; suppress only the paragraph strut.
        if (node.localName === "p" || node.closest('[data-office-text-break="true"]')) { node.style.fontSize = "0px"; node.style.lineHeight = "0px"; node.style.minHeight = "0px"; }
        else { node.style.fontSize = size * scale + "px"; if (Number.isFinite(line)) node.style.lineHeight = line * scale * (1 - (properties?.lineSpacingReduction ?? 0)) + "px"; }
      } };
      let scale = properties?.fontScale ?? 1;
      const measured = () => ({ width: Math.max(text.scrollWidth, text.clientWidth), height: Math.max(text.scrollHeight, text.clientHeight),
        availableWidth: (item.geometry.textRect?.width ?? 0) / .75, availableHeight: (item.geometry.textRect?.height ?? 0) / .75 });
      // A CSS border box cannot shrink below its padding. Compare against the
      // declared text rectangle so that oversized insets do not appear to fit.
      const overflow = (): boolean => { const size = measured(); return !!item.geometry.textRect && (size.height > size.availableHeight + 1 || size.width > size.availableWidth + 1); };
      apply(scale);
      if (properties?.autoFit === "normal") for (let pass = 0; pass < maxPasses && overflow(); pass++) {
        const size = measured(), next = Math.max(.01, scale * Math.min(.95, Math.sqrt(size.availableHeight / Math.max(1, size.height)), size.availableWidth / Math.max(1, size.width)));
        if (next >= scale) break; scale = next; apply(scale);
      }
      if (properties?.autoFit === "shape") for (let pass = 0; pass < maxPasses && overflow(); pass++) {
        const size = measured(); item.resize(item.width + Math.max(0, size.width - size.availableWidth) * .75, item.height + Math.max(0, size.height - size.availableHeight) * .75);
      }
      const exceeded = overflow();
      text.style.justifyContent = item.anchor === "ctr" ? "center" : item.anchor === "b" ? "flex-end" : item.anchor === "just" ? "space-between" : item.anchor === "dist" ? "space-around" : "flex-start";
      item.element.remove(); return Object.freeze({ element: item.element, text: item.input.text, widthPt: item.width, heightPt: item.height, geometry: item.geometry, scale, overflow: exceeded, nodes: item.nodes });
    });
    complete = true;
    return Object.freeze({ objects: Object.freeze(objects), fonts, usedFallbackLineMetrics });
  } finally {
    if (!complete) for (const input of bodies) input.text.element.remove();
    root.remove(); root.replaceChildren();
  }
}
