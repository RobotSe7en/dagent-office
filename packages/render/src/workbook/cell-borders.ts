import { DocumentError } from "@dagent-office/core";
import type { WorkbookBorderPaint } from "./border-style.js";
import { isWorkbookBorderKind } from "./border-style.js";
import { resolvedWorkbookColor, workbookDisplayColor } from "./color.js";

export interface WorkbookCellBorders {
  /** Logical column edges; native left/right mean leading/trailing. */
  readonly before?: WorkbookBorderPaint; readonly after?: WorkbookBorderPaint;
  readonly top?: WorkbookBorderPaint; readonly bottom?: WorkbookBorderPaint;
}
export interface WorkbookCellDiagonals { readonly up?: WorkbookBorderPaint; readonly down?: WorkbookBorderPaint }
export function workbookCellDiagonals(document: Document, style: CSSStyleDeclaration): WorkbookCellDiagonals {
  const enabled = (name: string): boolean => {
    const value = style.getPropertyValue(`--office-cell-border-diagonal-${name}`).trim();
    if (value && value !== "true" && value !== "false") throw new DocumentError("InvalidDocument", "A worksheet diagonal direction flag is invalid.");
    return value === "true";
  };
  const up = enabled("up"), down = enabled("down"), kind = style.getPropertyValue("--office-cell-border-diagonal-kind").trim();
  if (kind && !isWorkbookBorderKind(kind)) throw new DocumentError("InvalidDocument", "A worksheet diagonal border kind is invalid.");
  if (!kind || kind === "none" || !up && !down) return Object.freeze({});
  const paint = Object.freeze({ kind, color: resolvedWorkbookColor(document, style, "--office-cell-border-diagonal") ?? workbookDisplayColor(document, style.color) }) as WorkbookBorderPaint;
  return Object.freeze({ ...(up ? { up: paint } : {}), ...(down ? { down: paint } : {}) });
}

/** Read one resolved source cell, before shared-edge/merge composition. Native
 * column sides mirror with sheet columns; ordinary CSS left/right stay physical.
 * Contradictory native aliases are explicit failures, not guessed precedence. */
export function workbookCellBorders(document: Document, style: CSSStyleDeclaration, rightToLeft: boolean): WorkbookCellBorders {
  const native = (side: string): WorkbookBorderPaint | undefined => {
    const prefix = `--office-cell-border-${side}`, kind = style.getPropertyValue(prefix + "-kind").trim();
    if (!kind) return undefined;
    if (!isWorkbookBorderKind(kind)) throw new DocumentError("InvalidDocument", "A native worksheet border kind is invalid.");
    return Object.freeze({ kind, color: kind === "none" ? "rgba(0,0,0,0)" : resolvedWorkbookColor(document, style, prefix) ?? workbookDisplayColor(document, style.color) });
  };
  const css = (side: string): WorkbookBorderPaint | undefined => {
    const kind = style.getPropertyValue(`border-${side}-style`), rawWidth = style.getPropertyValue(`border-${side}-width`), width = Number.parseFloat(rawWidth) * .75;
    if (kind === "none" || kind === "hidden" || width === 0) return undefined;
    if (!["solid", "double", "dashed", "dotted"].includes(kind)) throw new DocumentError("UnsupportedFeature", "This CSS worksheet border style is not implemented.");
    if (!rawWidth.endsWith("px") || !Number.isFinite(width) || width <= 0) throw new DocumentError("InvalidDocument", "A worksheet CSS border has no valid computed width.");
    return Object.freeze({ kind: "css", style: kind as "solid" | "double" | "dashed" | "dotted", widthPt: width,
      color: workbookDisplayColor(document, style.getPropertyValue(`border-${side}-color`)) });
  };
  const column = (side: "left" | "right", alias: "start" | "end"): WorkbookBorderPaint | undefined => {
    const original = native(side), logical = native(alias);
    if (logical && original && original.kind !== "none" && (logical.kind !== original.kind || logical.color !== original.color)) {
      throw new DocumentError("UnsupportedFeature", "Conflicting native logical and legacy worksheet borders have no established precedence.");
    }
    const border = logical ?? original ?? css(rightToLeft ? side === "left" ? "right" : "left" : side);
    return border?.kind === "none" ? undefined : border;
  };
  const top = native("top") ?? css("top"), bottom = native("bottom") ?? css("bottom"), before = column("left", "start"), after = column("right", "end");
  return Object.freeze({ ...(before ? { before } : {}), ...(after ? { after } : {}), ...(top && top.kind !== "none" ? { top } : {}), ...(bottom && bottom.kind !== "none" ? { bottom } : {}) });
}
