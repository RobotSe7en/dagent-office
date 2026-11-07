import type { WorkbookCellInput } from "../contracts/workbook-operations.js";
import type { WorkbookManifest } from "../contracts/document.js";
import type { WorkbookChunkDescriptor, WorkbookRange } from "../contracts/workbook.js";
import type { DocumentReadLimits, WorkbookCell, WorkbookModel } from "../runtime.js";
import { assetDigest } from "../document/assets.js";
import { decodeText, encode } from "../document/package.js";
import { checkLimit, invalid } from "../document/validation.js";
import { CHUNK_COLUMNS, CHUNK_ROWS, decodeWorkbookCell, indexWorkbookChunk } from "../workbook/chunk-index.js";
import { parseFormula } from "../workbook/formula.js";
import { checkRange } from "../workbook/ranges.js";
import { workbookSource } from "../workbook/source.js";

export function escapeWorkbookText(value: string): string {
  for (const character of value) {
    const point = character.codePointAt(0)!;
    if (point >= 0xd800 && point <= 0xdfff || point === 0) invalid("Workbook text must contain well-formed Unicode without null characters.");
  }
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/\r/g, "&#13;");
}
export function rangeContains(range: WorkbookRange, row: number, column: number): boolean {
  return row >= range.from.row && row <= range.to.row && column >= range.from.column && column <= range.to.column;
}
export function rangesIntersect(a: WorkbookRange, b: WorkbookRange): boolean {
  return a.from.row <= b.to.row && a.to.row >= b.from.row && a.from.column <= b.to.column && a.to.column >= b.from.column;
}
interface EditedChunk { descriptor: WorkbookChunkDescriptor; cells: Map<number, string> }
const coordinate = (row: number, column: number): number => row * 16_384 + column;

/** A preparation owns only touched, bounded cell blocks. Other bytes stay shared. */
export class WorkbookCellEdit {
  private readonly chunks = new Map<string, EditedChunk>();
  private readonly originals: ReadonlyMap<string, import("../workbook/chunk-index.js").ChunkIndex>;
  constructor(private readonly base: WorkbookModel, readonly manifest: WorkbookManifest,
    readonly files: Map<string, Uint8Array>, private readonly limits: DocumentReadLimits, private readonly allocate: () => string) {
    this.originals = workbookSource(base)!.chunks;
  }
  private key(sheetId: string, row: number, column: number): string { return JSON.stringify([sheetId, Math.floor(row / CHUNK_ROWS), Math.floor(column / CHUNK_COLUMNS)]); }
  private chunk(sheetId: string, row: number, column: number): EditedChunk {
    const key = this.key(sheetId, row, column), known = this.chunks.get(key); if (known) return known;
    const position = { sheetId, row: Math.floor(row / CHUNK_ROWS) * CHUNK_ROWS, column: Math.floor(column / CHUNK_COLUMNS) * CHUNK_COLUMNS };
    let descriptor = this.manifest.chunks.find(chunk => chunk.sheetId === sheetId && chunk.row === position.row && chunk.column === position.column);
    if (!descriptor) {
      let path: string; do { path = "content/" + this.allocate() + ".html"; } while (this.files.has(path));
      // File quotas apply to the final candidate, independent of operation
      // order. Intermediate descriptors are bounded by the transaction and
      // node budget; flush plus package validation checks final file counts.
      checkLimit(this.chunks.size + 1, this.limits.maxNodes);
      descriptor = { ...position, path, sha256: "0".repeat(64) }; this.manifest.chunks.push(descriptor);
    }
    const index = this.originals.get(descriptor.path), cells = new Map<number, string>();
    for (const cell of index?.readWindow({ ...position, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }) ?? []) cells.set(coordinate(cell.row, cell.column), cell.html);
    const result = { descriptor, cells }; this.chunks.set(key, result); return result;
  }
  get(sheetId: string, row: number, column: number): WorkbookCell | undefined {
    const chunk = this.chunks.get(this.key(sheetId, row, column));
    if (!chunk) return this.base.getCell(sheetId, row, column);
    const html = chunk.cells.get(coordinate(row, column)); return html === undefined ? undefined : { sheetId, row, column, html, ...decodeWorkbookCell(html) };
  }
  *entries(sheetId?: string, formulasOnly = false): Generator<WorkbookCell> {
    for (const descriptor of [...this.manifest.chunks]) {
      if (sheetId !== undefined && descriptor.sheetId !== sheetId) continue;
      const edited = this.chunks.get(this.key(descriptor.sheetId, descriptor.row, descriptor.column));
      if (edited) {
        for (const [key, html] of [...edited.cells]) {
          const cell = { sheetId: descriptor.sheetId, row: Math.floor(key / 16_384), column: key % 16_384, html, ...decodeWorkbookCell(html) };
          if (!formulasOnly || cell.formula) yield cell;
        }
      } else if (formulasOnly) {
        for (const location of this.originals.get(descriptor.path)?.formulaLocations() ?? []) yield this.base.getCell(descriptor.sheetId, location.row, location.column)!;
      } else for (const cell of this.originals.get(descriptor.path)?.readWindow({ ...descriptor, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS }) ?? []) yield { sheetId: descriptor.sheetId, ...cell, ...decodeWorkbookCell(cell.html) };
    }
  }
  removeSheet(sheetId: string): void {
    for (const descriptor of this.manifest.chunks.filter(chunk => chunk.sheetId === sheetId)) {
      this.files.delete(descriptor.path); this.chunks.delete(this.key(sheetId, descriptor.row, descriptor.column));
    }
    this.manifest.chunks = this.manifest.chunks.filter(chunk => chunk.sheetId !== sheetId);
  }
  set(sheetId: string, input: WorkbookCellInput): void {
    const previous = this.get(sheetId, input.row, input.column);
    const cssClass = input.cssClass === undefined ? previous?.cssClass : input.cssClass;
    const formatId = input.formatId === undefined ? previous?.formatId : input.formatId;
    const sameStyle = (cssClass ?? undefined) === previous?.cssClass && (formatId ?? undefined) === previous?.formatId;
    if (!previous && !cssClass && !formatId && input.source.kind === "value" && input.source.value.type === "blank") return;
    if (previous && sameStyle && (input.source.kind === "value" && !previous.formula && JSON.stringify(input.source.value) === JSON.stringify(previous.value)
      || input.source.kind === "formula" && previous.formula?.expression === input.source.expression)) return;
    const attrs: Record<string, string> = { "data-column": String(input.column) };
    let html: string;
    if (input.source.kind === "formula") {
      parseFormula(input.source.expression, input);
      attrs["data-value-type"] = "blank"; attrs["data-formula"] = input.source.expression; attrs["data-cache-state"] = "stale"; html = "";
    } else if (input.source.kind === "richText") { attrs["data-value-type"] = "text"; html = input.source.html; }
    else {
      const value = input.source.value; attrs["data-value-type"] = value.type;
      if (value.type !== "blank" && value.type !== "text") attrs["data-value"] = String(value.value);
      html = value.type === "blank" ? "" : escapeWorkbookText(value.type === "boolean" ? value.value ? "TRUE" : "FALSE" : String(value.value)).replace(/\n/g, "<br>");
    }
    if (cssClass) attrs.class = cssClass; if (formatId) attrs["data-format-id"] = formatId;
    const source = "<td" + Object.entries(attrs).map(([name, value]) => " " + name + '="' + escapeWorkbookText(value) + '"').join("") + ">" + html + "</td>";
    const chunk = this.chunk(sheetId, input.row, input.column);
    // Validate even intermediate cells that a later operation would clear.
    const probe = this.serialize(chunk.descriptor, new Map([[coordinate(input.row, input.column), source]]));
    indexWorkbookChunk(probe, chunk.descriptor, this.limits.maxFileBytes, { maxNodes: this.limits.maxNodes, maxDepth: this.limits.maxDepth });
    if (attrs["data-value-type"] === "blank" && !attrs["data-formula"] && !cssClass && !formatId) chunk.cells.delete(coordinate(input.row, input.column));
    else chunk.cells.set(coordinate(input.row, input.column), source);
  }
  clear(sheetId: string, ranges: readonly WorkbookRange[], mode: "contents" | "all"): void {
    for (const range of ranges) checkRange(range);
    for (const descriptor of [...this.manifest.chunks]) {
      if (descriptor.sheetId !== sheetId || !ranges.some(range => rangesIntersect(range, { from: descriptor, to: { row: descriptor.row + CHUNK_ROWS - 1, column: descriptor.column + CHUNK_COLUMNS - 1 } }))) continue;
      const chunk = this.chunk(sheetId, descriptor.row, descriptor.column);
      for (const key of [...chunk.cells.keys()]) {
        const row = Math.floor(key / 16_384), column = key % 16_384;
        if (!ranges.some(range => rangeContains(range, row, column))) continue;
        if (mode === "all") chunk.cells.delete(key);
        else this.set(sheetId, { row, column, source: { kind: "value", value: { type: "blank" } } });
      }
    }
  }
  private serialize(descriptor: WorkbookChunkDescriptor, cells: ReadonlyMap<number, string>): Uint8Array {
    let html = '<!doctype html><html><head></head><body><table data-sheet-id="' + escapeWorkbookText(descriptor.sheetId) + '" data-row="' + descriptor.row + '" data-column="' + descriptor.column + '"><tbody>';
    let previous = -1;
    for (const [key, cell] of [...cells].sort((a, b) => a[0] - b[0])) {
      const row = Math.floor(key / 16_384);
      if (row !== previous) { if (previous >= 0) html += "</tr>"; html += '<tr data-row="' + row + '">'; previous = row; }
      html += cell; checkLimit(html.length, this.limits.maxFileBytes);
    }
    if (previous >= 0) html += "</tr>";
    const bytes = encode(html + "</tbody></table></body></html>"); checkLimit(bytes.byteLength, this.limits.maxFileBytes); return bytes;
  }
  flush(): void {
    for (const chunk of this.chunks.values()) {
      if (!chunk.cells.size) { this.files.delete(chunk.descriptor.path); this.manifest.chunks = this.manifest.chunks.filter(item => item.path !== chunk.descriptor.path); continue; }
      const bytes = this.serialize(chunk.descriptor, chunk.cells), previous = this.files.get(chunk.descriptor.path);
      // Preserve original bytes when the final cell set is unchanged, including source syntax.
      const original = this.originals.get(chunk.descriptor.path);
      const originalCells = original?.readWindow({ ...chunk.descriptor, rows: CHUNK_ROWS, columns: CHUNK_COLUMNS });
      if (previous && originalCells?.length === chunk.cells.size && originalCells.every(cell => chunk.cells.get(coordinate(cell.row, cell.column)) === cell.html)) continue;
      if (previous && decodeText(previous) === decodeText(bytes)) continue;
      this.files.set(chunk.descriptor.path, bytes); chunk.descriptor.sha256 = assetDigest(bytes);
    }
  }
}
