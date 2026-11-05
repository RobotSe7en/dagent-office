import { DocumentError } from "@dagent-office/core";
import type { WorkbookValue } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import type { WorkbookTextView } from "./text-view.js";
import { WorkbookTextDirection } from "./text-direction.js";
import { workbookReadingOrder, workbookTextAlignment } from "./text-alignment.js";
import { stackWorkbookText } from "./stacked-text.js";
import { measuredFontLineHeight } from "../font-line-height.js";

export interface WorkbookTextLayoutInput {
  readonly source: WorkbookSourceCell;
  readonly text: WorkbookTextView;
  readonly valueType: WorkbookValue["type"];
  readonly widthPt: number;
  readonly heightPt: number;
}
export interface WorkbookTextLayout {
  readonly element: HTMLElement;
  readonly text: WorkbookTextView;
  /** Unclipped height, including the vertical cell gutter, for row measurement. */
  readonly requiredHeightPt: number;
  readonly scale: number;
}
export interface WorkbookTextLayouts {
  readonly cells: readonly WorkbookTextLayout[];
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
}

/** Batch-resolve local fonts before measuring, once for the whole window. Input
 * text views are private detached projections; source nodes are never changed.
 * Return full-cell geometry for later pane clipping. This does not yet compose
 * neighboring-cell overflow or choose formatted numeric/date display strings. */
export function layoutWorkbookTexts(scope: HTMLElement, inputs: readonly WorkbookTextLayoutInput[], normalSpaceWidthPx: number, maxNodes = 50_000): WorkbookTextLayouts {
  const document = scope.ownerDocument, view = document.defaultView;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isFinite(normalSpaceWidthPx) || normalSpaceWidthPx < 0
    || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Worksheet text layout requires a connected surface, settled fonts and valid measurement limits.");
  let nodes = 3;
  if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet text layout exceeds its node budget.");
  const seen = new Set<HTMLElement>();
  for (const input of inputs) {
    if (input.source.element.ownerDocument !== document || !input.source.element.isConnected || input.text.element.ownerDocument !== document || input.text.element.parentNode || seen.has(input.text.element)
      || !Number.isFinite(input.widthPt / .75) || input.widthPt <= 0 || !Number.isFinite(input.heightPt / .75) || input.heightPt <= 0) throw new DocumentError("InvalidOperation", "Worksheet text layout requires distinct detached views and positive finite cell geometry.");
    seen.add(input.text.element); nodes += input.text.nodes + 1;
    if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet text layout exceeds its node budget.");
  }
  const root = markViewNode(document.createElement("div")); root.dataset.officeView = "workbook-text-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden;direction:ltr";
  const direction = new WorkbookTextDirection(scope);
  let complete = false;
  try {
    const prepared = inputs.map(input => {
      const style = view.getComputedStyle(input.source.element), readingOrder = workbookReadingOrder(style);
      // Native typed displays keep their own ordering: a worksheet RTL setting
      // must not turn -123 into 123- or #DIV/0! into !DIV/0#. Text values still
      // honor explicit reading order, including punctuation and mixed scripts.
      const textDirection = direction.read(input.text.element.textContent ?? "", input.valueType === "text" ? readingOrder : "context");
      const alignment = workbookTextAlignment(style, input.valueType, textDirection), element = markViewNode(document.createElement("div"));
      const textView = alignment.stacked ? stackWorkbookText(input.text, maxNodes - nodes + input.text.nodes) : input.text;
      nodes += textView.nodes - input.text.nodes;
      element.style.cssText = `position:relative;overflow:hidden;direction:ltr;width:${input.widthPt}pt;height:${input.heightPt}pt`;
      // Gutter is part of cell content geometry, independent of border stroke
      // width. Borders are a separate shared-edge layer, not CSS box borders.
      const indent = alignment.indentSpaces * normalSpaceWidthPx, left = 1.5 + (alignment.horizontal === "left" ? indent : 0), right = 1.5 + (alignment.horizontal === "right" ? indent : 0);
      if (!Number.isFinite(indent)) throw new DocumentError("InvalidOperation", "Worksheet indentation exceeds its numeric range.");
      const available = Math.max(0, input.widthPt / .75 - left - right), text = textView.element;
      Object.assign(text.style, { display: "block", position: "absolute", left: "0", top: "0", margin: "0", padding: "0", border: "0", boxSizing: "content-box", direction: textDirection,
        // Spreadsheet reference advances do not apply pair kerning. Explicit
        // OpenType control also covers engines that keep kerning despite the
        // high-level property on an inherited inline run.
        fontKerning: "none", fontFeatureSettings: '"kern" 0',
        unicodeBidi: "isolate", whiteSpace: alignment.wrap ? "pre-wrap" : "pre", overflowWrap: alignment.wrap ? "anywhere" : "normal", wordBreak: "normal",
        textAlign: alignment.stacked ? "center" : alignment.horizontal, width: alignment.wrap && !alignment.stacked ? `${available}px` : "max-content", height: "auto", transformOrigin: "0 0" });
      element.append(text); root.append(element); return { input, textView, alignment, element, available, left };
    });
    scope.append(root);
    const fonts = resolvePreviewFontsInScope(document, root, "[data-office-text-style],[data-office-script]");
    let usedFallbackLineMetrics = false;
    if (prepared.some(item => item.alignment.stacked)) {
      if (++nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Worksheet font measurement exceeds its node budget.");
      const context = document.createElement("canvas").getContext("2d");
      for (const item of prepared.filter(item => item.alignment.stacked)) {
        const styles = [item.textView.element, ...item.textView.element.querySelectorAll<HTMLElement>("[data-office-text-style],[data-office-script]")].map(node => {
          const computed = view.getComputedStyle(node);
          return { node, normal: computed.lineHeight === "normal", height: measuredFontLineHeight(context, computed) };
        });
        for (const { node, normal, height } of styles) if (normal) {
          if (height === undefined) usedFallbackLineMetrics = true;
          else node.style.lineHeight = `${height}px`;
        }
      }
    }
    const cells = prepared.map(({ input, textView, alignment, element, available, left }): WorkbookTextLayout => {
      const text = textView.element, measured = view.getComputedStyle(text), width = Number.parseFloat(measured.width), height = Number.parseFloat(measured.height);
      if (![width, height].every(value => Number.isFinite(value) && value >= 0)) throw new DocumentError("UnsupportedFeature", "This browser cannot measure worksheet text geometry.");
      const angle = alignment.rotationDegrees * Math.PI / 180, cosine = Math.cos(angle), sine = Math.sin(angle);
      const rotatedWidth = Math.abs(width * cosine) + Math.abs(height * sine), rotatedHeight = Math.abs(width * sine) + Math.abs(height * cosine);
      const minimumX = Math.min(0, width * cosine, -height * sine, width * cosine - height * sine), minimumY = Math.min(0, width * sine, height * cosine, width * sine + height * cosine);
      const scale = alignment.shrink && !alignment.wrap && !alignment.stacked && !/[\r\n]/.test(text.textContent ?? "") && rotatedWidth > 0 ? Math.min(1, available / rotatedWidth) : 1;
      const extra = available - rotatedWidth * scale, x = left + (alignment.horizontal === "center" ? extra / 2 : alignment.horizontal === "right" ? extra : 0) - minimumX * scale;
      const remaining = input.heightPt / .75 - 2 - rotatedHeight * scale;
      const y = 1 + (alignment.vertical === "center" ? remaining / 2 : alignment.vertical === "bottom" ? remaining : 0);
      text.style.left = `${x}px`; text.style.top = `${y - minimumY * scale}px`;
      text.style.transform = alignment.rotationDegrees ? `rotate(${alignment.rotationDegrees}deg) scale(${scale})` : scale === 1 ? "none" : `scale(${scale})`;
      return Object.freeze({ element, text: textView, requiredHeightPt: (rotatedHeight * scale + 2) * .75, scale });
    });
    for (const cell of cells) cell.element.remove();
    complete = true;
    return Object.freeze({ cells: Object.freeze(cells), fonts, usedFallbackLineMetrics });
  } finally {
    if (!complete) for (const input of inputs) input.text.element.remove();
    direction.dispose(); root.remove(); root.replaceChildren();
  }
}
