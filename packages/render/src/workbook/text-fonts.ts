import { DocumentError } from "@dagent-office/core";
import { officeFontSegments } from "../font-scripts.js";
import type { OfficeFontRole } from "../font-scripts.js";
import type { WorkbookSourceCell } from "./cell-source.js";

export interface WorkbookTextFont {
  readonly family: string; readonly sizePx: number; readonly weight: string; readonly style: string;
  readonly variant: string; readonly letterSpacing: string; readonly wordSpacing: string;
}
export interface WorkbookTextFontRun {
  readonly from: number; readonly to: number; readonly text: string;
  readonly role: OfficeFontRole; readonly font: WorkbookTextFont;
}

/** Read source CSS before adding any derived script wrappers. This keeps
 * source selectors and run inheritance intact while producing exact source
 * offsets for width measurement and later view projection. No font is loaded
 * or substituted here; the view owns local-font diagnostics and fallback. */
export function workbookTextFonts(source: WorkbookSourceCell, maxRuns = 50_000): readonly WorkbookTextFontRun[] {
  return workbookRichTextFonts(source.element, maxRuns);
}

/** Shared cell/paragraph font projection. A paragraph caller starts a new
 * script context, while adjacent spans keep combining/variation sequences. */
export function workbookRichTextFonts(element: HTMLElement, maxRuns = 50_000): readonly WorkbookTextFontRun[] {
  const view = element.ownerDocument.defaultView;
  if (!view || !element.isConnected || !Number.isSafeInteger(maxRuns) || maxRuns < 1) throw new DocumentError("InvalidOperation", "Workbook text fonts require a connected source projection and a positive run budget.");
  const runs: WorkbookTextFontRun[] = [], styles = new WeakMap<Element, CSSStyleDeclaration>(), fonts = new WeakMap<Element, Map<OfficeFontRole, WorkbookTextFont>>();
  let offset = 0, previous: OfficeFontRole = "ascii";
  const append = (text: string, parent: Element): void => {
    let computed = styles.get(parent); if (!computed) { computed = view.getComputedStyle(parent); styles.set(parent, computed); }
    let variants = fonts.get(parent); if (!variants) { variants = new Map(); fonts.set(parent, variants); }
    for (const segment of officeFontSegments(text, previous)) {
      if (runs.length >= maxRuns) throw new DocumentError("ResourceLimitExceeded", "Workbook text font runs exceed their budget.");
      let font = variants.get(segment.role);
      if (!font) {
        const family = computed.getPropertyValue(`--office-font-${segment.role}`).trim() || computed.getPropertyValue(`--office-minor-${segment.role}`).trim() || computed.fontFamily;
        if (!view.CSS.supports("font-family", family)) throw new DocumentError("InvalidDocument", "Workbook text has an invalid source font family.");
        const sizePx = Number.parseFloat(computed.fontSize);
        if (!Number.isFinite(sizePx) || sizePx < 0) throw new DocumentError("InvalidDocument", "Workbook text has an invalid computed font size.");
        font = Object.freeze({ family, sizePx, weight: computed.fontWeight, style: computed.fontStyle, variant: computed.fontVariant, letterSpacing: computed.letterSpacing, wordSpacing: computed.wordSpacing });
        variants.set(segment.role, font);
      }
      runs.push(Object.freeze({ from: offset + segment.from, to: offset + segment.to, text: segment.text, role: segment.role, font }));
      previous = segment.role;
    }
    offset += text.length;
  };
  const walker = element.ownerDocument.createTreeWalker(element, view.NodeFilter.SHOW_ELEMENT | view.NodeFilter.SHOW_TEXT);
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    if (node.nodeType === 3) append(node.textContent ?? "", node.parentElement!);
    else if ((node as Element).localName === "br") { append("\n", node as Element); previous = "ascii"; }
  }
  return Object.freeze(runs);
}
