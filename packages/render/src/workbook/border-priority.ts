import { DocumentError } from "@dagent-office/core";

const strength = {
  none: 0, hair: 1, thin: 2, dashed: 2, dotted: 2, dashDot: 2, dashDotDot: 2,
  medium: 3, mediumDashed: 3, mediumDashDot: 3, mediumDashDotDot: 3, slantDashDot: 3, double: 4, thick: 5,
} as const;
export type WorkbookBorderKind = keyof typeof strength;

/** Resolve two regular source borders sharing one edge. First precedes second
 * in source row/column order. At equal strength, physical bottom/right wins;
 * a mirrored column axis therefore reverses the tie. This follows the pinned
 * Calc reference, not CSS border-collapse rules. CF/merge-edge composition and
 * stroke geometry are separate; choosing slantDashDot does not render it. */
export function selectWorkbookBorder<First extends { readonly kind: WorkbookBorderKind }, Second extends { readonly kind: WorkbookBorderKind }>(first: First | undefined, second: Second | undefined, axis: "row" | "column", rightToLeft: boolean): First | Second | undefined {
  const rank = (border: { readonly kind: WorkbookBorderKind } | undefined): number => {
    if (!border) return 0;
    if (!Object.hasOwn(strength, border.kind)) throw new DocumentError("UnsupportedFeature", "The worksheet border kind is not supported.");
    return strength[border.kind];
  };
  const a = rank(first), b = rank(second);
  if (!a && !b) return undefined;
  return a > b || a === b && axis === "column" && rightToLeft ? first : second;
}
