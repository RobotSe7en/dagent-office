import { DocumentError } from "@dagent-office/core";
import type { WorkbookValue } from "@dagent-office/core";
import type { WorkbookReadingOrder } from "./text-direction.js";

export interface WorkbookTextAlignment {
  readonly horizontal: "left" | "center" | "right";
  readonly vertical: "top" | "center" | "bottom";
  readonly indentSpaces: number;
  readonly wrap: boolean;
  readonly shrink: boolean;
  readonly rotationDegrees: number;
  readonly stacked: boolean;
}
export function workbookTextGutters(alignment: WorkbookTextAlignment, normalSpaceWidthPx: number): { readonly left: number; readonly right: number } {
  const indent = alignment.indentSpaces * normalSpaceWidthPx;
  if (!Number.isFinite(normalSpaceWidthPx) || normalSpaceWidthPx < 0 || !Number.isFinite(indent)) throw new DocumentError("InvalidOperation", "Worksheet indentation requires finite nonnegative font metrics.");
  return { left: 1.5 + (alignment.horizontal === "left" ? indent : 0), right: 1.5 + (alignment.horizontal === "right" ? indent : 0) };
}
const invalid = (): never => { throw new DocumentError("InvalidDocument", "Worksheet text alignment has an invalid value."); };
function choice<T extends string>(style: CSSStyleDeclaration, name: string, values: readonly T[], fallback: T): T {
  const value = style.getPropertyValue(`--office-cell-${name}`).trim();
  if (!value) return fallback;
  return values.find(item => item === value) ?? invalid();
}
function integer(style: CSSStyleDeclaration, name: string, min: number, max: number): number {
  const value = style.getPropertyValue(`--office-cell-${name}`).trim();
  if (!value) return 0;
  if (!/^[+-]?\d+$/.test(value)) invalid();
  const number = Number(value); if (number < min || number > max) invalid(); return number;
}
export function workbookReadingOrder(style: CSSStyleDeclaration): WorkbookReadingOrder {
  return choice(style, "reading-order", ["context", "ltr", "rtl"], "context");
}
export function workbookTextRotation(style: CSSStyleDeclaration): Pick<WorkbookTextAlignment, "rotationDegrees" | "stacked"> {
  const rotation = integer(style, "text-rotation", 0, 255);
  if (rotation > 180 && rotation !== 255) invalid();
  return Object.freeze({ rotationDegrees: rotation === 255 ? 0 : rotation <= 90 ? -rotation : rotation - 90, stacked: rotation === 255 });
}

/** Resolve physical text alignment independently of worksheet column direction.
 * General uses the value type, never the appearance of its formatted string.
 * OOXML centers booleans/errors even though the pinned Calc reference does not. */
export function workbookTextAlignment(style: CSSStyleDeclaration, type: WorkbookValue["type"], direction: "ltr" | "rtl"): WorkbookTextAlignment {
  if (!["blank", "text", "number", "boolean", "error"].includes(type) || !["ltr", "rtl"].includes(direction)) invalid();
  const native = choice(style, "horizontal", ["general", "left", "center", "right", "fill", "justify", "centerContinuous", "distributed"], "general");
  let horizontal: string = native;
  if (!style.getPropertyValue("--office-cell-horizontal").trim()) {
    const css = style.textAlign;
    if (css === "center" || css === "left" || css === "right" || css === "justify") horizontal = css;
    else if (css === "end") horizontal = direction === "rtl" ? "left" : "right";
  }
  if (horizontal === "general") horizontal = type === "number" ? "right" : type === "boolean" || type === "error" ? "center" : direction === "rtl" ? "right" : "left";
  let vertical: string = choice(style, "vertical", ["top", "center", "bottom", "justify", "distributed"], "bottom");
  if (!style.getPropertyValue("--office-cell-vertical").trim()) {
    if (style.verticalAlign === "middle") vertical = "center";
    else if (style.verticalAlign === "top" || style.verticalAlign === "bottom") vertical = style.verticalAlign;
  }
  if (!["left", "center", "right"].includes(horizontal) || !["top", "center", "bottom"].includes(vertical)) {
    throw new DocumentError("UnsupportedFeature", "This worksheet alignment requires additional line or neighboring-cell layout.");
  }
  const indent = integer(style, "indent", 0, 255), relative = integer(style, "relative-indent", -15, 15);
  choice(style, "justify-last-line", ["true", "false"], "false");
  const wrap = choice(style, "wrap", ["true", "false"], ["pre-wrap", "pre-line", "break-spaces"].includes(style.whiteSpace) ? "true" : "false") === "true";
  const shrink = choice(style, "shrink-to-fit", ["true", "false"], "false") === "true";
  return Object.freeze({ horizontal: horizontal as WorkbookTextAlignment["horizontal"], vertical: vertical as WorkbookTextAlignment["vertical"],
    indentSpaces: horizontal === "center" ? 0 : Math.max(0, indent * 3 + relative), wrap, shrink, ...workbookTextRotation(style) });
}
