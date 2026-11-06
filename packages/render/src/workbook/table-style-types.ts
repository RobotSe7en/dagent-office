import type { WorkbookTableStyle } from "@dagent-office/core";

export type WorkbookTableRegion = WorkbookTableStyle["elements"][number]["region"];
export interface WorkbookBuiltinTableColor { readonly theme: number; readonly tint: number }
export interface WorkbookBuiltinTableDifferential {
  readonly font?: { readonly bold?: true; readonly color?: WorkbookBuiltinTableColor };
  readonly fill?: WorkbookBuiltinTableColor;
  /** Inner borders are region boundaries, not an instruction to outline each
   * cell. The region compositor chooses the applicable side for a coordinate. */
  readonly borders?: Readonly<Partial<Record<"left" | "right" | "top" | "bottom" | "horizontal" | "vertical",
    { readonly style: "none" | "thin" | "medium" | "thick" | "double"; readonly color?: WorkbookBuiltinTableColor }>>>;
}
export interface WorkbookBuiltinTableElement {
  readonly region: WorkbookTableRegion;
  readonly size: number;
  readonly differential: number;
}
