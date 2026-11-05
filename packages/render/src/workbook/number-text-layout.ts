import { DocumentError } from "@dagent-office/core";
import type { WorkbookDisplayFormat, WorkbookNumberDisplay, WorkbookNumberFormatter } from "@dagent-office/core";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { measureNumberDisplay } from "./number-layout.js";
import { createWorkbookNumberText } from "./number-text.js";
import { workbookTextAlignment, workbookTextGutters } from "./text-alignment.js";
import { layoutWorkbookTexts } from "./text-layout.js";
import type { WorkbookTextLayoutInput, WorkbookTextLayouts } from "./text-layout.js";
import { workbookAsciiWidth } from "./ascii-width.js";

export interface WorkbookNumberLayoutInput {
  readonly source: WorkbookSourceCell;
  readonly value: number;
  readonly format?: WorkbookDisplayFormat;
  readonly widthPt: number;
  readonly heightPt: number;
  readonly bordered?: boolean;
}
type Overflow = Extract<WorkbookNumberDisplay, { kind: "overflow" }>["reason"];
export interface WorkbookNumberLayouts extends WorkbookTextLayouts {
  readonly overflow: readonly (Overflow | undefined)[];
  readonly locales: readonly string[];
}

/** Resolve fonts for derived scalar displays, measure full shaped candidates,
 * then use the shared cell alignment/rotation layout. Temporary glyph probes
 * and failed windows never mutate source HTML or the current published grid. */
export function layoutWorkbookNumbers(scope: HTMLElement, inputs: readonly WorkbookNumberLayoutInput[], formatter: WorkbookNumberFormatter,
  normalSpaceWidthPx: number, maxNodes = 50_000, maxTextUnits: number = 8 * 1024 * 1024): WorkbookNumberLayouts {
  const document = scope.ownerDocument, view = document.defaultView;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isFinite(normalSpaceWidthPx) || normalSpaceWidthPx < 0
    || !Number.isSafeInteger(maxNodes) || maxNodes < 1 || !Number.isSafeInteger(maxTextUnits) || maxTextUnits < 1) throw new DocumentError("InvalidOperation", "Numeric layout requires a connected surface, settled fonts and positive budgets.");
  const root = markViewNode(document.createElement("div")); root.dataset.officeView = "workbook-number-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden;direction:ltr";
  let liveNodes = 1, initialUnits = 0, outputUnits = 0;
  const remainingNodes = (): number => {
    if (liveNodes >= maxNodes) throw new DocumentError("ResourceLimitExceeded", "Numeric layout exceeds its node budget.");
    return maxNodes - liveNodes;
  };
  const finals: WorkbookTextLayoutInput[] = [], overflow: (Overflow | undefined)[] = [], locales: string[] = [];
  try {
    const prepared = inputs.map(input => {
      if (input.source.element.ownerDocument !== document || !input.source.element.isConnected || !Number.isFinite(input.value)
        || !Number.isFinite(input.widthPt / .75) || input.widthPt <= 0 || !Number.isFinite(input.heightPt / .75) || input.heightPt <= 0) throw new DocumentError("InvalidOperation", "Numeric layout requires finite values, matching connected sources and positive geometry.");
      const display = formatter.format(input.value, input.format);
      const parts = display.kind === "parts" ? display.parts.map(part => ({ kind: "text" as const, text: part.kind === "text" ? part.text : part.character })) : [{ kind: "text" as const, text: "#" }];
      initialUnits += parts.reduce((sum, part) => sum + part.text.length, 0);
      if (initialUnits > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Numeric measurement exceeds its cumulative text budget.");
      const text = createWorkbookNumberText(input.source, parts, display.color, remainingNodes()); liveNodes += text.nodes;
      root.append(text.element); return { input, text };
    });
    scope.append(root);
    resolvePreviewFontsInScope(document, root, "[data-office-text-style],[data-office-script]");
    const context = document.createElement("canvas").getContext("2d");
    for (const { input, text } of prepared) {
      const alignment = workbookTextAlignment(view.getComputedStyle(input.source.element), "number", "ltr"), gutter = workbookTextGutters(alignment, normalSpaceWidthPx);
      const available = Math.max(0, input.widthPt / .75 - gutter.left - gutter.right), fit = !alignment.shrink && !alignment.wrap && !alignment.rotationDegrees && !alignment.stacked;
      const families = new Map<string, string>(), widths = new Map<string, number>();
      for (const span of text.element.querySelectorAll<HTMLElement>("[data-office-script]")) families.set(span.dataset.officeScript!, view.getComputedStyle(span).fontFamily);
      const general = !input.format || "builtinId" in input.format && input.format.builtinId === 0 || "code" in input.format && input.format.code === "General";
      const ascii = text.element.querySelector<HTMLElement>('[data-office-script="ascii"]');
      const fast = general && ascii ? workbookAsciiWidth(context, view.getComputedStyle(ascii)) : undefined;
      const measure = (value: string): number => {
        const known = widths.get(value); if (known !== undefined) return known;
        const quick = fast?.(value);
        // Preserve the exact DOM decision at a fit boundary, including layout
        // engine subpixel quantization. Complex format parts never use this.
        if (quick !== undefined && Math.abs(quick - available) > 1 / 16) return quick;
        const sample = createWorkbookNumberText(input.source, [{ kind: "text", text: value }], undefined, remainingNodes());
        sample.element.style.cssText += ';display:inline-block;position:absolute;width:max-content;white-space:pre;font-kerning:none;font-feature-settings:"kern" 0;direction:ltr;unicode-bidi:isolate';
        root.append(sample.element);
        try {
          let missing = false;
          for (const span of sample.element.querySelectorAll<HTMLElement>("[data-office-script]")) {
            const family = families.get(span.dataset.officeScript!);
            if (family) span.style.fontFamily = family; else missing = true;
          }
          if (missing) {
            resolvePreviewFontsInScope(document, sample.element, "[data-office-script]");
            for (const span of sample.element.querySelectorAll<HTMLElement>("[data-office-script]")) families.set(span.dataset.officeScript!, view.getComputedStyle(span).fontFamily);
          }
          const width = sample.element.getBoundingClientRect().width;
          if (!Number.isFinite(width) || width < 0) throw new DocumentError("UnsupportedFeature", "This browser cannot measure numeric display text.");
          if (widths.size >= 64) widths.delete(widths.keys().next().value!);
          widths.set(value, width);
          return width;
        } finally { sample.element.remove(); }
      };
      const display = formatter.format(input.value, input.format, fit ? { availableWidth: available, measure } : undefined);
      let reason: Overflow | undefined = display.kind === "overflow" ? display.reason : undefined;
      let measured = display.kind === "parts" ? measureNumberDisplay(display.parts, available, measure) : undefined;
      if (fit && measured?.overflow) reason = "width";
      if (reason) {
        measured = measureNumberDisplay([{ kind: "fill", character: "#" }], available, measure);
        if (!measured.parts.some(part => part.kind === "text" && part.text)) { const width = measure("#"); measured = { parts: [{ kind: "text", text: "#", width }], width, overflow: true }; }
      }
      const parts = measured!.parts;
      outputUnits += parts.reduce((sum, part) => sum + (part.kind === "text" ? part.text.length : 0), 0);
      if (outputUnits > maxTextUnits) throw new DocumentError("ResourceLimitExceeded", "Numeric display exceeds its cumulative text budget.");
      text.element.remove(); liveNodes -= text.nodes;
      const final = createWorkbookNumberText(input.source, parts, display.color, remainingNodes()); liveNodes += final.nodes;
      finals.push({ source: input.source, text: final, valueType: "number", widthPt: input.widthPt, heightPt: input.heightPt,
        ...(input.bordered === undefined ? {} : { bordered: input.bordered }) }); overflow.push(reason); locales.push(display.locale);
    }
  } finally { root.remove(); root.replaceChildren(); }
  const layout = layoutWorkbookTexts(scope, finals, normalSpaceWidthPx, maxNodes, maxTextUnits);
  return Object.freeze({ ...layout, overflow: Object.freeze(overflow), locales: Object.freeze(locales) });
}
