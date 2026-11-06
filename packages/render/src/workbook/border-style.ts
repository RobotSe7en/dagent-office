import { DocumentError } from "@doffi/core";
const kinds = ["none", "thin", "medium", "dashed", "dotted", "thick", "double", "hair", "mediumDashed", "dashDot", "mediumDashDot", "dashDotDot", "mediumDashDotDot", "slantDashDot"] as const;
export type WorkbookBorderKind = typeof kinds[number];
const kindSet = new Set<string>(kinds);
export function isWorkbookBorderKind(value: string): value is WorkbookBorderKind { return kindSet.has(value); }

export type WorkbookBorderStyle = { readonly kind: WorkbookBorderKind }
  | { readonly kind: "css"; readonly style: "solid" | "double" | "dashed" | "dotted"; readonly widthPt: number };
export type WorkbookBorderPaint = WorkbookBorderStyle & { readonly color: string };
export interface WorkbookBorderStroke { readonly width: number; readonly offset: number; readonly dash?: readonly number[]; readonly cap?: "butt" | "round"; readonly phasePt?: number }
const profiles: Readonly<Partial<Record<WorkbookBorderKind, readonly WorkbookBorderStroke[]>>> = {
  none: [], hair: [{ width: .1, offset: 0 }], thin: [{ width: .75, offset: 0 }], medium: [{ width: 1.75, offset: 0 }], thick: [{ width: 2.5, offset: 0 }],
  double: [{ width: .5, offset: -.625 }, { width: .5, offset: .625 }],
  dashed: [{ width: .75, offset: 0, dash: [3, 1] }], dotted: [{ width: .75, offset: 0, dash: [.5, 1] }],
  mediumDashed: [{ width: 1.75, offset: 0, dash: [8, 2.5] }],
  dashDot: [{ width: .75, offset: 0, dash: [8, 2.5, 2.5, 2.5] }], mediumDashDot: [{ width: 1.75, offset: 0, dash: [8, 2.5, 2.5, 2.5] }],
  dashDotDot: [{ width: .75, offset: 0, dash: [8, 2.5, 2.5, 2.5, 2.5, 2.5] }], mediumDashDotDot: [{ width: 1.75, offset: 0, dash: [8, 2.5, 2.5, 2.5, 2.5, 2.5] }],
  // Two staggered lanes retain the slanted dash ends illustrated in ECMA-376
  // §18.18.3, rather than treating this style as an ordinary dashed line.
  slantDashDot: [{ width: .75, offset: -.375, dash: [8.25, .75, 3.75, .75] }, { width: .75, offset: .375, dash: [7.5, 1.5, 3, 1.5], phasePt: .75 }],
};

/** CSS point widths remain exact; they are not rounded to a native named style.
 * Native profiles retain the independently observed stroke widths and spacing.
 * CSS double divides its width into two strokes and a gap of equal thickness. */
export function workbookBorderStrokes(border: WorkbookBorderStyle): readonly WorkbookBorderStroke[] {
  if (border.kind !== "css") {
    if (!Object.hasOwn(profiles, border.kind)) throw new DocumentError("UnsupportedFeature", "This worksheet border has no verified stroke geometry.");
    return profiles[border.kind]!;
  }
  const width = border.widthPt;
  if (!Number.isFinite(width * 3) || width / 3 <= 0 || !["solid", "double", "dashed", "dotted"].includes(border.style)) throw new DocumentError("InvalidDocument", "A CSS worksheet border requires a supported style and positive finite point width.");
  if (border.style === "double") return [{ width: width / 3, offset: -width / 3 }, { width: width / 3, offset: width / 3 }];
  if (border.style === "dashed") return [{ width, offset: 0, dash: [width * 3, width] }];
  if (border.style === "dotted") return [{ width, offset: 0, dash: [0, width * 2], cap: "round" }];
  return [{ width, offset: 0 }];
}
