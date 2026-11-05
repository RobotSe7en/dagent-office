import { DocumentError } from "@dagent-office/core";
import type { WorkbookNumberFormatter, WorkbookSnapshot } from "@dagent-office/core";
import type { FontResolution } from "../contracts.js";
import { resolvePreviewFontsInScope } from "../fonts.js";
import { copyWorkbookTextStyle } from "./text-style.js";
import { workbookComputedTextFont } from "./text-fonts.js";
import { layoutWorkbookCellTexts } from "./cell-text-layout.js";
import type { WorkbookCellTextInput } from "./cell-text-layout.js";
import { workbookRowFontKey } from "./row-metric-cache.js";

export interface WorkbookRowTextLayouts {
  readonly heights: readonly (number | undefined)[];
  readonly fonts: readonly FontResolution[];
  readonly usedFallbackLineMetrics: boolean;
}
/** Empty cells contribute their font's normal line box, never the rotation,
 * wrapping or number-format literals of an invented cell value. The M is a
 * private font probe, like the column-width digits, and has no source offsets. */
export function layoutWorkbookRowTexts(scope: HTMLElement, snapshot: WorkbookSnapshot, inputs: readonly WorkbookCellTextInput[], formatter: WorkbookNumberFormatter,
  normalSpaceWidthPx: number, fontCache: Map<string, number>, fontKey: (input: WorkbookCellTextInput) => string = workbookRowFontKey): WorkbookRowTextLayouts {
  if (!inputs.length) return Object.freeze({ heights: Object.freeze([]), fonts: Object.freeze([]), usedFallbackLineMetrics: false });
  const document = scope.ownerDocument, view = document.defaultView;
  if (!scope.isConnected || !view || document.fonts.status === "loading" || !Number.isFinite(normalSpaceWidthPx) || normalSpaceWidthPx < 0
    || inputs.some(input => !input.source.element.isConnected || input.source.element.ownerDocument !== document))
    throw new DocumentError("InvalidOperation", "Row font measurement requires connected sources and settled local fonts.");
  const normal = inputs.filter(input => input.value && input.value.type !== "blank");
  const measured = normal.length ? layoutWorkbookCellTexts(scope, snapshot, normal, formatter, { normalSpaceWidthPx }, 50_000, 8 * 1024 * 1024) : { cells: [], fonts: [], usedFallbackLineMetrics: false };
  const heights = new Map<WorkbookCellTextInput, number | undefined>(normal.map((input, index) => [input, measured.cells[index]?.text?.requiredHeightPt]));
  const keys = inputs.map(fontKey), pending = new Map<string, WorkbookCellTextInput>();
  for (const [index, input] of inputs.entries()) if (!fontCache.has(keys[index]!)) pending.set(keys[index]!, input);
  if (!pending.size) return Object.freeze({ heights: Object.freeze(inputs.map((input, index) => Math.max(heights.get(input) ?? 0, fontCache.get(keys[index]!)!))),
    fonts: Object.freeze([...measured.fonts]), usedFallbackLineMetrics: measured.usedFallbackLineMetrics });
  const root = document.createElement("div"); root.dataset.officeView = "workbook-row-font-measurement"; root.setAttribute("aria-hidden", "true");
  root.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0;width:0;height:0;overflow:hidden";
  if (pending.size * 2 + 1 > 50_000) throw new DocumentError("ResourceLimitExceeded", "Row font probes exceed their temporary node budget.");
  try {
    for (const input of pending.values()) {
      const probe = document.createElement("span"), computed = copyWorkbookTextStyle(input.source.element, probe);
      probe.style.fontFamily = workbookComputedTextFont(computed, "ascii", view).family;
      probe.style.cssText += ";display:inline-block;position:absolute;white-space:pre;margin:0;padding:0;border:0;width:max-content;height:auto;text-decoration:none";
      probe.textContent = "M"; root.append(probe);
    }
    scope.append(root);
    const fonts = pending.size ? resolvePreviewFontsInScope(document, root, "span") : [], measuredFonts = new Map<string, number>();
    for (const [index, key] of [...pending.keys()].entries()) {
      const height = (root.children[index] as HTMLElement).getBoundingClientRect().height;
      if (!Number.isFinite(height) || height < 0) throw new DocumentError("UnsupportedFeature", "This browser cannot measure empty-cell font leading.");
      const required = (height + 2) * .75; measuredFonts.set(key, required);
      if (fontCache.size < 1024 && key.length <= 4096) fontCache.set(key, required);
    }
    for (const [index, input] of inputs.entries()) heights.set(input, Math.max(heights.get(input) ?? 0, fontCache.get(keys[index]!) ?? measuredFonts.get(keys[index]!)!));
    return Object.freeze({ heights: Object.freeze(inputs.map(input => heights.get(input))), fonts: Object.freeze([...measured.fonts, ...fonts]), usedFallbackLineMetrics: measured.usedFallbackLineMetrics });
  } finally { root.remove(); root.replaceChildren(); }
}
