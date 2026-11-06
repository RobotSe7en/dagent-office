import { sha256 } from "@noble/hashes/sha2.js";
import { bytesToHex } from "@noble/hashes/utils.js";
import { DocumentError } from "@dagent-office/core";
import type { DocumentReadLimits, WorkbookChunkDescriptor, WorkbookCoordinate } from "@dagent-office/core";
import { encode, textHtml } from "./values.js";

/** UTF-8 accounting before concatenating or encoding generated source. Input is
 * already validated Unicode or generated ASCII markup. */
export function utf8Bytes(source: string): number {
  let bytes = 0;
  for (let index = 0; index < source.length; index++) {
    const unit = source.charCodeAt(index);
    if (unit < 0x80) bytes++; else if (unit < 0x800) bytes += 2;
    else if (unit >= 0xd800 && unit <= 0xdbff) { bytes += 4; index++; } else bytes += 3;
  }
  return bytes;
}
interface PendingChunk { readonly prefix: string; readonly parts: string[]; bytes: number; row: number }
const suffix = "</tbody></table></body></html>";

/** Owns one generated package's cumulative budget across every worksheet. */
export class WorkbookChunks {
  readonly files: Map<string, Uint8Array> = new Map();
  readonly descriptors: WorkbookChunkDescriptor[] = [];
  private usedBytes = 0;
  private readonly pending = new Map<number, PendingChunk>();
  private band = -1;
  private sheetId = "";
  private previous = -1;
  constructor(private readonly limits: DocumentReadLimits) {}
  private account(bytes: number): void {
    if (bytes + this.usedBytes > this.limits.maxPackageBytes) throw new DocumentError("ResourceLimitExceeded", "Generated workbook source exceeds its package budget.");
    this.usedBytes += bytes;
  }
  addFile(path: string, bytes: Uint8Array): void {
    if (this.files.has(path)) throw new DocumentError("InvalidDocument", "A generated workbook part is duplicated.");
    if (this.files.size + this.pending.size >= this.limits.maxFiles || bytes.length > this.limits.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "Generated workbook parts exceed their count or byte budget.");
    this.account(bytes.length); this.files.set(path, bytes);
  }
  startSheet(sheetId: string): void { this.flush(); this.sheetId = sheetId; this.band = -1; this.previous = -1; }
  append(coordinate: WorkbookCoordinate, html: string): void {
    const { row, column } = coordinate, position = row * 16384 + column;
    if (!this.sheetId || !Number.isSafeInteger(row) || !Number.isSafeInteger(column) || row < 0 || row > 1_048_575 || column < 0 || column > 16383 || position <= this.previous) throw new DocumentError("InvalidDocument", "Generated workbook coordinates must be ordered within a worksheet.");
    this.previous = position;
    const band = Math.floor(row / 256) * 256, columnBand = Math.floor(column / 32) * 32;
    if (band !== this.band) { this.flush(); this.band = band; }
    let chunk = this.pending.get(columnBand);
    if (!chunk) {
      if (this.files.size + this.pending.size >= this.limits.maxFiles) throw new DocumentError("ResourceLimitExceeded", "Generated workbook parts exceed their count budget.");
      const prefix = `<!doctype html><html><head></head><body><table data-sheet-id="${textHtml(this.sheetId)}" data-row="${band}" data-column="${columnBand}"><tbody>`;
      const bytes = utf8Bytes(prefix) + suffix.length;
      this.account(bytes); chunk = { prefix, parts: [], bytes, row: -1 }; this.pending.set(columnBand, chunk);
    }
    const before = chunk.row === row ? "" : (chunk.row >= 0 ? "</tr>" : "") + `<tr data-row="${row}">`;
    // Each row's final closing tag is reserved on first entry, including the
    // last row which is only emitted while flushing.
    const added = utf8Bytes(html) + before.length + (chunk.row === row || chunk.row >= 0 ? 0 : 5);
    if (chunk.bytes + added > this.limits.maxFileBytes) throw new DocumentError("ResourceLimitExceeded", "A generated workbook chunk exceeds its byte budget.");
    this.account(added); chunk.bytes += added; chunk.parts.push(before, html); chunk.row = row;
  }
  finish(): void { this.flush(); this.sheetId = ""; }
  private flush(): void {
    for (const [column, chunk] of [...this.pending].sort(([a], [b]) => a - b)) {
      const bytes = encode(chunk.prefix + chunk.parts.join("") + "</tr>" + suffix), path = `sheets/${this.sheetId}/${this.band}-${column}.html`;
      if (bytes.length !== chunk.bytes) throw new DocumentError("InvalidDocument", "Generated chunk byte accounting is inconsistent.");
      if (this.files.has(path)) throw new DocumentError("InvalidDocument", "A generated chunk is duplicated.");
      this.files.set(path, bytes); this.descriptors.push({ sheetId: this.sheetId, row: this.band, column, path, sha256: bytesToHex(sha256(bytes)) });
    }
    this.pending.clear();
  }
}
