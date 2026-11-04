import { DocumentError } from "@dagent-office/core";
import { resolvedWorkbookColor } from "./color.js";

/** Resolve one already-cascaded cell fill. Native fill declarations retain
 * their own none/solid distinction; ordinary package CSS can use background
 * color directly. Pattern/gradient geometry must not silently become solid. */
export function workbookCellFill(document: Document, style: CSSStyleDeclaration): string {
  const kind = style.getPropertyValue("--office-cell-fill-kind").trim();
  if (!kind) return style.backgroundColor;
  if (kind === "none") return "transparent";
  if (kind !== "solid") throw new DocumentError("UnsupportedFeature", "This worksheet fill requires unsupported pattern or gradient geometry.");
  const color = resolvedWorkbookColor(document, style, "--office-cell-fill");
  if (!color) throw new DocumentError("InvalidDocument", "A solid worksheet fill has no color.");
  return color;
}
