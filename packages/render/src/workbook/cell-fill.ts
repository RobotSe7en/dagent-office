import { DocumentError } from "@dagent-office/core";
import { resolvedWorkbookColor, workbookDisplayColor } from "./color.js";
import { isWorkbookFillPattern } from "./fill-pattern.js";
import type { WorkbookFillPattern } from "./fill-pattern.js";

export interface WorkbookGradientStop { readonly position: number; readonly color: string }
export type WorkbookCellFill =
  | { readonly kind: "solid"; readonly color: string }
  | { readonly kind: "pattern"; readonly pattern: WorkbookFillPattern; readonly foreground: string; readonly background: string }
  | { readonly kind: "linear"; readonly degree: number; readonly stops: readonly WorkbookGradientStop[] }
  | { readonly kind: "path"; readonly left: number; readonly right: number; readonly top: number; readonly bottom: number; readonly stops: readonly WorkbookGradientStop[] };

/** Resolve source colors and bounded parameters before geometry is created.
 * The returned paint is derived, immutable and independent of source CSS. */
export function workbookCellFill(document: Document, style: CSSStyleDeclaration, maxStops = 4096): WorkbookCellFill {
  if (!Number.isSafeInteger(maxStops) || maxStops < 1) throw new DocumentError("InvalidOperation", "A cell fill needs a positive gradient-stop budget.");
  const kind = style.getPropertyValue("--office-cell-fill-kind").trim();
  if (!kind || kind === "none") return Object.freeze({ kind: "solid", color: kind ? "transparent" : style.backgroundColor });
  if (kind === "gradient") {
    const number = (name: string, fallback?: number): number => {
      const raw = style.getPropertyValue("--office-cell-gradient-" + name).trim();
      if (!raw && fallback !== undefined) return fallback;
      if (!/^[+-]?(?:\d+(?:\.\d*)?|\.\d+)(?:e[+-]?\d+)?$/i.test(raw) || !Number.isFinite(Number(raw))) throw new DocumentError("InvalidDocument", "A worksheet gradient parameter is not finite numeric data.");
      return Number(raw);
    };
    const count = number("stops"), type = style.getPropertyValue("--office-cell-gradient-kind").trim() || "linear";
    if (!Number.isSafeInteger(count) || count < 0 || !["linear", "path"].includes(type)) throw new DocumentError("InvalidDocument", "A worksheet gradient type or stop count is invalid.");
    if (count > maxStops) throw new DocumentError("ResourceLimitExceeded", "A worksheet gradient exceeds its stop budget.");
    const stops: WorkbookGradientStop[] = [];
    for (let index = 0; index < count; index++) {
      const position = number(`stop-${index}-position`), color = resolvedWorkbookColor(document, style, `--office-cell-gradient-stop-${index}`);
      if (position < 0 || position > 1 || color === undefined) throw new DocumentError("InvalidDocument", "A worksheet gradient stop needs a fractional position and color.");
      stops.push(Object.freeze({ position, color }));
    }
    // Stable sorting leaves equal-position hard transitions in source order.
    // Only this derived list is sorted; persisted declarations remain intact.
    stops.sort((a, b) => a.position - b.position); Object.freeze(stops);
    const degree = number("degree", 0), sides = { left: number("left", 0), right: number("right", 0), top: number("top", 0), bottom: number("bottom", 0) };
    if (Object.values(sides).some(amount => amount < 0 || amount > 1)) throw new DocumentError("InvalidDocument", "A worksheet path gradient convergence is outside its fractional range.");
    return type === "linear" ? Object.freeze({ kind: "linear", degree, stops }) : Object.freeze({ kind: "path", ...sides, stops });
  }
  if (isWorkbookFillPattern(kind)) return Object.freeze({ kind: "pattern", pattern: kind,
    foreground: resolvedWorkbookColor(document, style, "--office-cell-fill") ?? workbookDisplayColor(document, style.getPropertyValue("--office-system-foreground").trim() || "#000000"),
    background: resolvedWorkbookColor(document, style, "--office-cell-fill-background") ?? workbookDisplayColor(document, style.getPropertyValue("--office-system-background").trim() || "#ffffff") });
  if (kind !== "solid") throw new DocumentError("InvalidDocument", "A worksheet fill has an unknown pattern kind.");
  const color = resolvedWorkbookColor(document, style, "--office-cell-fill");
  if (!color) throw new DocumentError("InvalidDocument", "A solid worksheet fill has no color.");
  return Object.freeze({ kind: "solid", color });
}
