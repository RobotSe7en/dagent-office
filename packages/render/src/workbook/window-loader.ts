import { DocumentError } from "@dagent-office/core";
import type { DeepReadonly, WorkbookCell, WorkbookSheet, WorkbookSnapshot } from "@dagent-office/core";
import type { AxisDimensions } from "./dimensions.js";
import { projectWorkbookCells, projectionWindows, WorkbookMergeIndex } from "./projection.js";
import type { WorkbookGridCell } from "./projection.js";
import { workbookViewport } from "./viewport.js";
import type { WorkbookViewport, WorkbookViewportLayout } from "./viewport.js";
import { cellSourceUnits, checkWindowIdentity, windowLimits } from "./window-source.js";
import type { WorkbookWindowLimits, WorkbookWindowSource } from "./window-source.js";

export interface WorkbookWindowFrame {
  readonly sheetId: string;
  readonly documentId: string; readonly revision: number;
  readonly layout: WorkbookViewportLayout;
  readonly projections: readonly WorkbookGridCell[];
  readonly storedCellCount: number;
  getCell(row: number, column: number): WorkbookCell | undefined;
}

/** At most one accepted viewport is retained. A complete frame replaces the old
 * frame atomically; cancellation, errors and late responses cannot splice old
 * data into new geometry. Every loader is bound to one immutable source. */
export class WorkbookWindowLoader {
  private controller: AbortController | undefined;
  private accepted: WorkbookWindowFrame | undefined;
  private readonly limits: WorkbookWindowLimits;
  private readonly snapshot: WorkbookSnapshot;
  private readonly sheets: ReadonlyMap<string, { readonly sheet: DeepReadonly<WorkbookSheet>; readonly merges: WorkbookMergeIndex }>;
  private disposed = false;
  constructor(private readonly source: WorkbookWindowSource, limits: Partial<WorkbookWindowLimits> = {}) {
    this.limits = windowLimits(limits); this.snapshot = source.snapshot;
    this.sheets = new Map(this.snapshot.metadata.workbook.sheets.map(sheet => [sheet.sheetId, { sheet, merges: new WorkbookMergeIndex(sheet.merges) }]));
  }
  get frame(): WorkbookWindowFrame | undefined { return this.accepted; }
  async load(sheetId: string, dimensions: { readonly rows: AxisDimensions; readonly columns: AxisDimensions }, viewport: WorkbookViewport): Promise<boolean> {
    if (this.disposed) throw new DocumentError("InvalidOperation", "This workbook window loader has been disposed.");
    this.controller?.abort(); const controller = new AbortController(); this.controller = controller;
    const entry = this.sheets.get(sheetId);
    if (!entry) throw new DocumentError("InvalidOperation", "The requested worksheet is not part of this workbook source.");
    const layout = workbookViewport(dimensions, entry.sheet.frozen, viewport, this.limits.maxCells);
    const projections = projectWorkbookCells(dimensions, layout, entry.merges, this.limits.maxCells);
    const windows = Object.freeze(projectionWindows(sheetId, projections).map(window => Object.freeze(window)));
    const { documentId, revision } = this.snapshot.manifest;
    try {
      const response = windows.length ? await this.source.read(Object.freeze({ documentId, revision, windows, ...this.limits }), controller.signal) : { documentId, revision, cells: [] };
      if (controller.signal.aborted || this.disposed || this.controller !== controller) return false;
      checkWindowIdentity(response, this.snapshot);
      if (response.cells.length > this.limits.maxCells) throw new DocumentError("ResourceLimitExceeded", "Workbook window response exceeds its cell budget.");
      const expected = new Set(projections.map(cell => cell.row * 16_384 + cell.column)), cells = new Map<number, WorkbookCell>(); let units = 0;
      for (const cell of response.cells) {
        const key = cell.row * 16_384 + cell.column;
        if (cell.sheetId !== sheetId || !Number.isSafeInteger(cell.row) || !Number.isSafeInteger(cell.column) || cell.row < 0 || cell.row >= 1_048_576 || cell.column < 0 || cell.column >= 16_384 || !expected.has(key) || cells.has(key)) throw new DocumentError("InvalidDocument", "Workbook window response contains duplicate or unrequested coordinates.");
        units += cellSourceUnits(cell);
        if (units > this.limits.maxCodeUnits) throw new DocumentError("ResourceLimitExceeded", "Workbook window response exceeds its source-text budget.");
        cells.set(key, Object.freeze({ ...cell, value: Object.freeze({ ...cell.value }), ...(cell.formula ? { formula: Object.freeze({ ...cell.formula }) } : {}) }));
      }
      for (const pane of layout.panes) {
        for (const cell of [...pane.rows, ...pane.columns]) Object.freeze(cell);
        Object.freeze(pane.rows); Object.freeze(pane.columns); Object.freeze(pane);
      }
      Object.freeze(layout.panes); Object.freeze(layout);
      for (const cell of projections) { Object.freeze(cell.range.from); Object.freeze(cell.range.to); Object.freeze(cell.range); Object.freeze(cell.rect); Object.freeze(cell.clip); Object.freeze(cell); }
      this.accepted = Object.freeze({ sheetId, documentId, revision, layout, projections: Object.freeze(projections), storedCellCount: cells.size,
        getCell: (row: number, column: number) => Number.isSafeInteger(row) && Number.isSafeInteger(column) && row >= 0 && row < 1_048_576 && column >= 0 && column < 16_384 ? cells.get(row * 16_384 + column) : undefined });
      return true;
    } catch (error) {
      if (controller.signal.aborted || this.disposed || this.controller !== controller) return false;
      throw error;
    }
  }
  dispose(): void { this.disposed = true; this.controller?.abort(); this.controller = undefined; this.accepted = undefined; }
}
