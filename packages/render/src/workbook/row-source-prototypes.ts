import { DocumentError } from "@doffi/core";
import type { DeepReadonly, WorkbookCell, WorkbookSheet, WorkbookSnapshot } from "@doffi/core";
import { WorkbookCellSources } from "./cell-source.js";
import type { WorkbookSourceCell } from "./cell-source.js";
import { workbookCellFormat } from "./cell-format.js";
import { checkedPlainWorkbookCell } from "./plain-cell-source.js";
import { cellSourceUnits } from "./window-source.js";

/** One automatic-height pass owns these connected typography prototypes. Only
 * plain cells under class-only CSS with no table/rule layers may share them.
 * Every cell still passes the complete attribute/value/identity check. */
export class WorkbookRowSourcePrototypes {
  private readonly values = new Map<string, { readonly sources: WorkbookCellSources; readonly source: WorkbookSourceCell }>();
  private units = 0;
  private retainedNodes = 0;
  private disposed = false;
  private readonly sheet: DeepReadonly<WorkbookSheet>;
  constructor(private readonly scope: HTMLElement, private readonly snapshot: WorkbookSnapshot, private readonly sheetId: string) {
    const sheet = snapshot.metadata.workbook.sheets.find(sheet => sheet.sheetId === sheetId);
    if (!sheet) throw new DocumentError("InvalidOperation", "Workbook typography prototypes require a source worksheet.");
    this.sheet = sheet;
  }
  get nodes(): number { return this.retainedNodes; }
  get(cell: WorkbookCell): WorkbookSourceCell | undefined {
    if (this.disposed) throw new DocumentError("InvalidOperation", "Workbook typography prototypes have been disposed.");
    if (!checkedPlainWorkbookCell(cell)) return undefined;
    const format = workbookCellFormat(this.sheet, cell, cell), key = JSON.stringify(format), known = this.values.get(key);
    if (known) return known.source;
    const units = cellSourceUnits(cell) + key.length;
    if (this.values.size >= 128 || this.units + units > 8 * 1024 * 1024) return undefined;
    const sources = new WorkbookCellSources(this.scope, this.snapshot, { documentId: this.snapshot.manifest.documentId, revision: this.snapshot.manifest.revision,
      sheetId: this.sheetId, sourceCoordinates: [{ row: cell.row, column: cell.column }], getCell: (row, column) => row === cell.row && column === cell.column ? cell : undefined });
    const source = sources.get(cell.row, cell.column)!;
    this.values.set(key, { sources, source }); this.units += units; this.retainedNodes += 5 + Number(cell.text.length > 0);
    return source;
  }
  dispose(): void { if (this.disposed) return; this.disposed = true; for (const value of this.values.values()) value.sources.dispose(); this.values.clear(); this.units = 0; this.retainedNodes = 0; }
}
