import { DEFAULT_WORKBOOK_INDEXED_COLORS, DocumentError } from "@dagent-office/core";
import type { WorkbookNumberFormatColor } from "@dagent-office/core";
import { officeFontSegments } from "../font-scripts.js";
import { markViewNode } from "../view-nodes.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { workbookDisplayColor } from "./color.js";
import { workbookComputedTextFont } from "./text-fonts.js";
import { copyWorkbookTextStyle, decorateWorkbookText } from "./text-style.js";
import type { WorkbookTextView } from "./text-view.js";

export type WorkbookNumberTextPart = { readonly kind: "text"; readonly text: string } | { readonly kind: "space"; readonly width: number };
const named = { black: "#000000", blue: "#0000ff", cyan: "#00ffff", green: "#00ff00", magenta: "#ff00ff", red: "#ff0000", white: "#ffffff", yellow: "#ffff00" };
function numberColor(document: Document, style: CSSStyleDeclaration, color: WorkbookNumberFormatColor): string {
  if (color.kind === "named") return workbookDisplayColor(document, named[color.name]);
  // Format Color1 is native palette entry 8, not entry 1. A custom palette
  // must provide every referenced color; do not fill its gaps from defaults.
  const index = color.index + 7, selected = style.getPropertyValue(`--office-indexed-${index}`).trim();
  if (selected) return workbookDisplayColor(document, selected);
  for (let item = 0; item < 64; item++) if (style.getPropertyValue(`--office-indexed-${item}`).trim()) throw new DocumentError("InvalidDocument", "A numeric format references a missing custom palette color.");
  const fallback = DEFAULT_WORKBOOK_INDEXED_COLORS[index];
  if (!fallback) throw new DocumentError("InvalidDocument", "A numeric format palette index is invalid.");
  return workbookDisplayColor(document, fallback);
}

/** Numeric/date strings and padding are derived displays. They deliberately
 * have no source text offsets; selections refer to their worksheet cell. */
export function createWorkbookNumberText(source: WorkbookSourceCell, parts: readonly WorkbookNumberTextPart[],
  color?: WorkbookNumberFormatColor, maxNodes = 50_000): WorkbookTextView {
  const document = source.element.ownerDocument, view = document.defaultView;
  if (!view || !source.element.isConnected || !Number.isSafeInteger(maxNodes) || maxNodes < 1) throw new DocumentError("InvalidOperation", "Numeric text requires a connected source and a positive node budget.");
  const root = markViewNode(document.createElement("span")), style = copyWorkbookTextStyle(source.element, root);
  root.style.whiteSpace = "pre"; if (color) root.style.color = numberColor(document, style, color);
  const underline = style.getPropertyValue("--office-font-underline").trim();
  if (underline && underline !== "none") root.dataset.officeUnderlinedText = "true";
  if (underline.endsWith("Accounting")) root.dataset.officeAccountingText = "true";
  let nodes = 1, units = 0;
  const account = (amount: number): void => { nodes += amount; if (nodes > maxNodes) throw new DocumentError("ResourceLimitExceeded", "Numeric text exceeds its node budget."); };
  for (const part of parts) {
    if (part.kind === "space") {
      if (!Number.isFinite(part.width) || part.width < 0) throw new DocumentError("InvalidOperation", "Numeric spacing requires a finite nonnegative width.");
      account(1); const span = document.createElement("span"); span.dataset.officeNumberSpace = "true";
      span.style.cssText = `display:inline-block;width:${part.width}px`; span.setAttribute("aria-hidden", "true"); root.append(span);
    } else {
      units += part.text.length;
      if (units > 32767) throw new DocumentError("ResourceLimitExceeded", "Numeric text exceeds its character budget.");
      for (const segment of officeFontSegments(part.text)) {
        account(2); const span = document.createElement("span"), text = document.createTextNode(segment.text);
        span.dataset.officeScript = segment.role; span.style.fontFamily = workbookComputedTextFont(style, segment.role, view).family;
        decorateWorkbookText(style, span, account).append(text); root.append(span);
      }
    }
  }
  return Object.freeze({ element: root, projections: Object.freeze([]), nodes });
}
