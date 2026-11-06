import type { WorkbookCell } from "../runtime.js";
import { freeze } from "../document/validation.js";
import { createUtf8TextCopy } from "../document/package.js";
import { decodeWorkbookCell } from "./chunk-index.js";
import type { IndexedCellSource } from "./chunk-index.js";

/** A single immutable model owns this bounded derived LRU. Indexed source
 * bytes remain authoritative; no parse tree or implicit blank is retained. */
export class WorkbookCellReadCache {
  private readonly entries = new Map<string, { readonly cell: WorkbookCell; readonly units: number }>();
  private units = 0;
  private readonly copyText = createUtf8TextCopy();
  constructor(private readonly maxEntries: number, private readonly maxUnits = 8 * 1024 * 1024) {}
  read(sheetId: string, source: IndexedCellSource): WorkbookCell {
    const key = JSON.stringify([sheetId, source.row, source.column]), known = this.entries.get(key);
    if (known) { this.entries.delete(key); this.entries.set(key, known); return known.cell; }
    // A substring may retain the entire decoded chunk in some engines. A UTF-8
    // copy detaches this already validated source slice before retaining it;
    // the decoder then also derives text/attributes from that small string.
    const html = this.copyText(source.html);
    const cell = freeze({ sheetId, row: source.row, column: source.column, html, ...decodeWorkbookCell(html) });
    const units = key.length + sheetId.length + cell.html.length + cell.text.length + (cell.value.type === "text" || cell.value.type === "error" ? cell.value.value.length : 0)
      + (cell.formula?.expression.length ?? 0) + (cell.cssClass?.length ?? 0) + (cell.formatId?.length ?? 0);
    if (units <= this.maxUnits && this.maxEntries > 0) {
      while (this.entries.size && (this.entries.size >= this.maxEntries || this.units + units > this.maxUnits)) {
        const first = this.entries.keys().next().value!, previous = this.entries.get(first)!; this.entries.delete(first); this.units -= previous.units;
      }
      this.entries.set(key, { cell, units }); this.units += units;
    }
    return cell;
  }
}
