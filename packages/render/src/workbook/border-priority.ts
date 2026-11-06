import { DocumentError } from "@dagent-office/core";
import type { WorkbookBorderStyle, WorkbookBorderKind } from "./border-style.js";
export type { WorkbookBorderKind } from "./border-style.js";

const strength: Readonly<Record<WorkbookBorderKind, number>> = {
  none: 0, hair: 1, thin: 2, dashed: 2, dotted: 2, dashDot: 2, dashDotDot: 2,
  medium: 3, mediumDashed: 3, mediumDashDot: 3, mediumDashDotDot: 3, slantDashDot: 3, double: 4, thick: 5,
} as const;

/** Resolve two regular source borders sharing one edge. First precedes second
 * in source row/column order. At equal strength, physical bottom/right wins;
 * a mirrored column axis therefore reverses the tie. This follows the pinned
 * Calc reference, not CSS border-collapse rules. CF/merge-edge composition and
 * stroke geometry are separate; choosing slantDashDot does not render it. */
export function selectWorkbookBorder<First extends WorkbookBorderStyle, Second extends WorkbookBorderStyle>(first: First | undefined, second: Second | undefined, axis: "row" | "column", rightToLeft: boolean): First | Second | undefined {
  const rank = (border: WorkbookBorderStyle | undefined): readonly [number, number] => {
    if (!border) return [0, 0];
    if (border.kind === "css") {
      if (!Number.isFinite(border.widthPt * 3) || border.widthPt <= 0 || !["solid", "double", "dashed", "dotted"].includes(border.style)) throw new DocumentError("InvalidDocument", "A CSS worksheet border has an invalid width or style.");
      return [border.widthPt, border.style === "double" ? 1 : 0];
    }
    if (!Object.hasOwn(strength, border.kind)) throw new DocumentError("UnsupportedFeature", "The worksheet border kind is not supported.");
    return [[0, .1, .75, 1.75, 1.75, 2.5][strength[border.kind]]!, border.kind === "double" ? 1 : 0];
  };
  const a = rank(first), b = rank(second);
  if (!a[0] && !b[0]) return undefined;
  const comparison = a[0] - b[0] || a[1] - b[1];
  return comparison > 0 || comparison === 0 && axis === "column" && rightToLeft ? first : second;
}
