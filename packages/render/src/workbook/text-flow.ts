import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookSheet, WorkbookValue } from "@dagent-office/core";
import type { WorkbookCellSources } from "./cell-source.js";
import type { AxisDimensions } from "./dimensions.js";
import { intersectRect, paneRect } from "./projection.js";
import type { WorkbookGridCell, WorkbookMergeIndex, WorkbookRect } from "./projection.js";
import type { WorkbookTextLayout } from "./text-layout.js";
import type { WorkbookWindowFrame } from "./window-loader.js";

/** Neighbor-aware display geometry. Values, formulas, merges and explicit
 * continuous styles remain source facts; overflow never invents a merge or
 * changes the grid's hit/selection coordinates. */
export class WorkbookTextFlow {
  private readonly complete: ReadonlySet<number>;
  private readonly known: ReadonlySet<number>;
  private readonly panes: ReadonlyMap<WorkbookGridCell["pane"], WorkbookWindowFrame["layout"]["panes"][number]>;
  private work = 0;
  private readonly appearances = new WeakMap<CSSStyleDeclaration, { rotation: string; horizontal: string; wrap: boolean; shrink: boolean }>();
  private appearance(row: number, column: number) {
    const style = this.style(row, column), known = this.appearances.get(style); if (known) return known;
    const value = { rotation: style.getPropertyValue("--office-cell-text-rotation").trim(), horizontal: style.getPropertyValue("--office-cell-horizontal").trim(),
      wrap: style.getPropertyValue("--office-cell-wrap").trim() === "true", shrink: style.getPropertyValue("--office-cell-shrink-to-fit").trim() === "true" };
    this.appearances.set(style, value); return value;
  }
  constructor(private readonly frame: WorkbookWindowFrame, private readonly sheet: DeepReadonly<WorkbookSheet>,
    private readonly dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, private readonly merges: WorkbookMergeIndex,
    private readonly sources: WorkbookCellSources, private readonly style: (row: number, column: number) => CSSStyleDeclaration, private readonly maxWork: number) {
    this.complete = new Set(frame.completeRows); this.known = new Set(frame.sourceCoordinates.map(point => point.row * 16_384 + point.column));
    this.panes = new Map(frame.layout.panes.map(pane => [pane.kind, pane]));
  }
  private step(): void { if (++this.work > this.maxWork) throw new DocumentError("ResourceLimitExceeded", "Worksheet text-flow discovery exceeds its geometry budget."); }
  private has(row: number, column: number): boolean { return this.complete.has(row) || this.known.has(row * 16_384 + column); }
  private physical(rect: WorkbookRect): WorkbookRect { return this.sheet.rightToLeft ? { ...rect, x: this.frame.layout.widthPt - rect.x - rect.width } : rect; }

  /** Return potential offscreen origins before measuring. Final ink/pane
   * intersection rejects distant origins without materializing display DOM. */
  candidates(value: (row: number, column: number) => DeepReadonly<WorkbookValue> | undefined): readonly WorkbookGridCell[] {
    const result: WorkbookGridCell[] = [], existing = new Set(this.frame.projections.map(cell => `${cell.pane}:${cell.row}:${cell.column}`));
    for (const point of this.frame.sourceCoordinates) {
      this.step(); const current = value(point.row, point.column);
      const { rotation, horizontal } = this.appearance(point.row, point.column);
      const continuous = horizontal === "centerContinuous";
      const rotated = Boolean(rotation && rotation !== "0" && rotation !== "90" && rotation !== "180" && rotation !== "255");
      if (current?.type !== "text" && !(continuous && current && current.type !== "blank") && !rotated) continue;
      const range = this.merges.at(point.row, point.column) ?? { from: point, to: point };
      if (range.from.row !== point.row || range.from.column !== point.column) continue;
      const x = this.dimensions.columns.position(range.from.column), y = this.dimensions.rows.position(range.from.row);
      const width = this.dimensions.columns.position(range.to.column + 1) - x, height = this.dimensions.rows.position(range.to.row + 1) - y;
      if (!width || !height) continue;
      for (const pane of this.frame.layout.panes) {
        this.step(); if (existing.has(`${pane.kind}:${point.row}:${point.column}`) || !pane.rows.some(row => row.index >= range.from.row && row.index <= range.to.row)) continue;
        const frozen = pane.kind === "corner" || pane.kind === "frozenColumns";
        if (frozen ? range.from.column >= this.sheet.frozen.columns : range.to.column < this.sheet.frozen.columns) continue;
        const rect = { x: x - pane.sourceLeftPt + pane.inlineStartPt, y: y - pane.sourceTopPt + pane.blockStartPt, width, height };
        const bounds = paneRect(pane), clip = intersectRect({ x: bounds.x, y: rect.y, width: bounds.width, height: rect.height }, bounds);
        if (clip) result.push({ pane: pane.kind, row: point.row, column: point.column, range, rect, clip });
      }
    }
    return result;
  }

  /** A centerContinuous span includes only explicit blank cells carrying the
   * same cell style. It stops at a merge, nonblank/formula or frozen partition. */
  textRect(projection: WorkbookGridCell): WorkbookRect {
    if (this.appearance(projection.row, projection.column).horizontal !== "centerContinuous"
      || projection.range.from.row !== projection.range.to.row || projection.range.from.column !== projection.range.to.column) return projection.rect;
    const original = this.sources.get(projection.row, projection.column)!, pane = this.panes.get(projection.pane)!;
    let column = projection.column;
    for (;;) {
      this.step(); const next = this.dimensions.columns.neighbor(column, 1); if (next === undefined || !this.has(projection.row, next)) break;
      const frozen = pane.kind === "corner" || pane.kind === "frozenColumns";
      if (frozen && next >= this.sheet.frozen.columns || !frozen && next < this.sheet.frozen.columns || this.merges.at(projection.row, next)) break;
      const cell = this.frame.getCell(projection.row, next), source = this.sources.get(projection.row, next);
      if (!cell || cell.formula || cell.value.type !== "blank" || !source || source.format.cssClass !== original.format.cssClass) break;
      column = next;
    }
    return { ...projection.rect, width: this.dimensions.columns.position(column + 1) - this.dimensions.columns.position(projection.column) };
  }

  clip(projection: WorkbookGridCell, textRect: WorkbookRect, text: WorkbookTextLayout, type: WorkbookValue["type"]): WorkbookRect | undefined {
    this.step(); const pane = this.panes.get(projection.pane)!, physicalPane = this.physical(paneRect(pane)), body = this.physical(textRect);
    const { rotation: rawRotation, horizontal, wrap, shrink } = this.appearance(projection.row, projection.column), rotation = rawRotation ? Number(rawRotation) : 0;
    const rotated = rotation !== 0 && rotation !== 255;
    const ordinary = type === "text" && !rotation && !wrap && !shrink && !["fill", "justify", "distributed", "centerContinuous"].includes(horizontal)
      && projection.range.from.row === projection.range.to.row && projection.range.from.column === projection.range.to.column;
    let left = body.x, right = body.x + body.width;
    const inkLeft = body.x + text.ink.xPt, inkRight = inkLeft + text.ink.widthPt;
    if (rotated) { left = Math.min(left, inkLeft); right = Math.max(right, inkRight); }
    else if (ordinary) {
      for (const direction of [-1, 1] as const) {
        const target = direction < 0 ? Math.max(physicalPane.x, inkLeft) : Math.min(physicalPane.x + physicalPane.width, inkRight);
        let boundary = direction < 0 ? left : right, column = projection.column;
        const step = (this.sheet.rightToLeft ? -direction : direction) as -1 | 1;
        while (direction < 0 ? target < boundary : target > boundary) {
          this.step(); const next = this.dimensions.columns.neighbor(column, step); if (next === undefined || !this.has(projection.row, next) || this.merges.at(projection.row, next)) break;
          if ((next < this.sheet.frozen.columns) !== (projection.column < this.sheet.frozen.columns)) break;
          const cell = this.frame.getCell(projection.row, next); if (cell && (cell.formula || cell.value.type !== "blank")) break;
          const logicalX = this.dimensions.columns.position(next) - pane.sourceLeftPt + pane.inlineStartPt;
          const physical = this.physical({ x: logicalX, y: body.y, width: this.dimensions.columns.size(next), height: body.height });
          boundary = direction < 0 ? physical.x : physical.x + physical.width; column = next;
        }
        if (direction < 0) left = Math.min(left, boundary); else right = Math.max(right, boundary);
      }
    }
    const clip = intersectRect({ x: left, y: body.y, width: right - left, height: body.height }, physicalPane);
    if (!clip) return undefined;
    if (ordinary || rotated) {
      text.element.style.overflow = "visible";
      text.element.style.clipPath = `inset(0 ${-Math.max(0, text.ink.xPt + text.ink.widthPt - body.width) / .75}px 0 ${Math.min(0, text.ink.xPt) / .75}px)`;
    }
    return this.sheet.rightToLeft ? { ...clip, x: this.frame.layout.widthPt - clip.x - clip.width } : clip;
  }
}
