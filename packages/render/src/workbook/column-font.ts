import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookMetadata } from "@dagent-office/core";
import { resolvePreviewFontsInScope } from "../fonts.js";
import type { FontResolution } from "../contracts.js";

export interface WorkbookColumnFont {
  readonly cssClass: string;
  readonly font: string;
  /** Maximum individual digit advance in CSS pixels at the source font size. */
  readonly maximumDigitWidthPx: number;
  /** Normal-style space advance, also used by native cell indentation. */
  readonly spaceWidthPx: number;
  readonly fonts: readonly FontResolution[];
}

/** Styles must already be installed in this surface and local fonts settled.
 * Source CSS is never changed: substitutions apply only to the temporary probe
 * and remain in its returned font diagnostics. No network fonts are requested. */
export function measureWorkbookColumnFont(scope: HTMLElement, workbook: DeepReadonly<WorkbookMetadata>, sheetId: string): WorkbookColumnFont {
  const document = scope.ownerDocument, sheet = workbook.sheets.find(sheet => sheet.sheetId === sheetId);
  if (!sheet || !scope.isConnected || !document.defaultView || document.fonts.status === "loading") throw new DocumentError("InvalidOperation", "Column font measurement requires a connected worksheet surface with settled fonts.");
  const cssClass = workbook.columnWidthFontCssClass ?? sheet.defaultCssClass;
  const wrapper = document.createElement("div"), probe = document.createElement("span");
  wrapper.style.cssText = "position:absolute;visibility:hidden;pointer-events:none;contain:layout;left:0;top:0";
  if (workbook.themeCssClass) wrapper.className = workbook.themeCssClass;
  probe.className = cssClass; probe.textContent = "0123456789"; wrapper.append(probe); scope.append(wrapper);
  try {
    const fonts = resolvePreviewFontsInScope(document, wrapper, "span"), style = document.defaultView.getComputedStyle(probe);
    const canvas = document.createElement("canvas"), context = canvas.getContext("2d");
    if (!context) throw new DocumentError("UnsupportedFeature", "This browser cannot measure workbook column fonts.");
    const font = `${style.fontStyle} ${style.fontWeight} ${style.fontSize} ${style.fontFamily}`;
    context.font = font;
    const maximumDigitWidthPx = Math.max(...[..."0123456789"].map(digit => context.measureText(digit).width));
    const spaceWidthPx = context.measureText(" ").width;
    if (!Number.isFinite(maximumDigitWidthPx) || maximumDigitWidthPx <= 0 || !Number.isFinite(spaceWidthPx) || spaceWidthPx < 0) throw new DocumentError("InvalidOperation", "The resolved workbook column font has invalid character advances.");
    return Object.freeze({ cssClass, font, maximumDigitWidthPx, spaceWidthPx, fonts });
  } finally { wrapper.remove(); }
}
